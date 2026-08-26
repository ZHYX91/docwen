"""ENEX → Markdown converter.

Parses Evernote Export (ENEX) XML files, extracts notes with their
attachments, and converts the note content (ENML wrapped HTML) to
Markdown using ``markdownify``.

The converter:
- Only writes to staging via ``WorkspaceHandle``.
- Checks cancellation before expensive operations.
- Reports progress through ``ProgressSink``.
- Returns a ``ConversionResult`` with ``ArtifactManifest`` entries.
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import os
import re
import time
import uuid
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from importlib.util import find_spec
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


@dataclass(frozen=True)
class _EnexResource:
    """A single resource (attachment) extracted from an ENEX note."""

    md5: str
    mime: str
    data: bytes
    file_name: str | None = None


class EnexToMarkdownConverter:
    """Convert an ENEX (Evernote Export) file to Markdown.

    Each ENEX note becomes one Markdown file.  Embedded resources
    (images, attachments) are extracted to the staging directory.
    """

    def convert(self, context: ConverterContext) -> Any:
        """Run the ENEX → Markdown conversion.

        Args:
            context: The plugin execution context.

        Returns:
            ``ConversionResult`` with staging artifacts.
        """
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionMetrics,
            ConversionResult,
        )

        t_start = time.monotonic()
        task_id = context.request.request_id
        input_path = context.workspace.input_path
        options = context.request.options or {}
        keep_images = bool(options.get("to_md_keep_images", True))
        enable_ocr = bool(options.get("to_md_enable_ocr", False))
        yaml_key_labels = options.get("yaml_key_labels")
        from docwen_core.export_semantics import get_markdown_export_modes, resolve_markdown_request_policy

        request_policy = resolve_markdown_request_policy(context)

        export_modes = get_markdown_export_modes(
            "markup",
            extraction_mode=options.get("image_mode"),
            ocr_placement_mode=options.get("ocr_placement"),
            semantics=request_policy.export,
        )
        image_mode = str(export_modes["image_extraction_mode"]).strip().lower()
        if image_mode not in {"file", "base64", "embed", "omit"}:
            image_mode = "file"
        ocr_placement = str(export_modes["ocr_placement_mode"]).strip().lower()
        if ocr_placement not in {"image_md", "main_md"}:
            ocr_placement = "image_md"
        ocr_language = str(options.get("ocr_language") or "auto")
        current_locale = str(options.get("locale") or "zh_CN")
        image_link_style = str(options.get("image_link_style") or "").strip() or None

        # 1. Check cancellation
        context.cancellation.check()

        # 2. Verify markdownify availability
        if find_spec("markdownify") is None:
            from docwen_core.models.result import ConversionErrorInfo

            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="dependency_missing",
                    message=("ENEX→MD requires the 'markdownify' library. Install it with: pip install markdownify"),
                    diagnostic_code="ENEX2MD-MISSING-DEP",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message="Missing required dependency: markdownify",
                        code="ENEX2MD-MISSING-DEP",
                    ),
                ],
            )

        # 3. Report start
        context.progress.report_progress(0.0, "Starting ENEX → Markdown conversion")
        context.logger.info(f"ENEX→MD: reading {input_path}")

        # 4. Parse the ENEX file
        context.cancellation.check()
        context.progress.report_progress(10.0, "Parsing ENEX structure...")

        aux_artifacts: list[Any] = []

        try:
            md_content, stats = self._parse_enex(
                input_path,
                context,
                aux_artifacts,
                keep_images=keep_images,
                enable_ocr=enable_ocr,
                image_mode=image_mode,
                ocr_placement=ocr_placement,
                ocr_language=ocr_language,
                current_locale=current_locale,
                image_link_style=image_link_style,
                yaml_key_labels=yaml_key_labels,
            )
        except Exception as exc:
            context.logger.error(f"ENEX→MD conversion failed: {exc}")
            from docwen_core.models.result import ConversionErrorInfo

            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="ENEX2MD-PARSE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Failed to parse ENEX: {exc}",
                        code="ENEX2MD-PARSE-ERROR",
                    ),
                ],
            )

        # 5. Write markdown to staging
        context.cancellation.check()
        context.progress.report_progress(80.0, "Writing Markdown to staging...")

        output_path = context.workspace.create_artifact_path("primary", ".md")
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md_content)
        except OSError as exc:
            context.logger.error(f"ENEX→MD write failed: {exc}")
            from docwen_core.models.result import ConversionErrorInfo

            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=f"Failed to write output file: {exc}",
                    diagnostic_code="ENEX2MD-WRITE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"File write error at {output_path}: {exc}",
                        code="ENEX2MD-WRITE-ERROR",
                    ),
                ],
            )

        # 6. Build artifact manifest
        input_basename = os.path.basename(input_path)
        suggested_name = input_basename.rsplit(".", 1)[0] + ".md"

        from docwen_core.models.artifact import ArtifactManifest

        artifact = ArtifactManifest(
            artifact_id=str(uuid.uuid4()),
            kind="primary",
            staging_path=output_path,
            suggested_name=suggested_name,
            media_type="text/markdown",
            metadata={
                "note_count": stats.get("notes", 0),
                "resource_count": stats.get("resources", 0),
            },
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)

        # 7. Report completion
        context.progress.report_artifact_ready(artifact.artifact_id, suggested_name)
        context.progress.report_progress(100.0, "Conversion complete")
        context.logger.info(f"ENEX→MD complete: {stats.get('notes', 0)} notes, {stats.get('resources', 0)} resources")

        duration_ms = (time.monotonic() - t_start) * 1000.0

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact, *aux_artifacts],
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=(f"Converted ENEX to Markdown: {stats['notes']} notes, {stats['resources']} resources"),
                    code="ENEX2MD-OK",
                ),
            ],
            error=None,
            metrics=ConversionMetrics(
                duration_ms=duration_ms,
                input_bytes=os.path.getsize(input_path) if os.path.isfile(input_path) else 0,
                output_bytes=len(md_content.encode("utf-8")),
                extra=stats,
            ),
        )

    # ── Internal parsing ──────────────────────────────────────────────

    def _parse_enex(
        self,
        input_path: str,
        context: ConverterContext,
        aux_artifacts: list[Any],
        *,
        keep_images: bool,
        enable_ocr: bool,
        image_mode: str,
        ocr_placement: str,
        ocr_language: str,
        current_locale: str,
        image_link_style: str | None,
        yaml_key_labels: object | None,
    ) -> tuple[str, dict[str, int]]:
        """Parse ENEX file into Markdown.

        Args:
            input_path: Path to the ENEX file.
            context: Plugin execution context.
            aux_artifacts: Output list — auxiliary artifacts are appended here.

        Returns:
            (markdown_text, stats_dict) where stats_dict has keys
            'notes' (int) and 'resources' (int).
        """
        from markdownify import markdownify as md_convert

        tree = ET.parse(input_path)
        root = tree.getroot()

        notes: list[str] = []
        note_titles: list[str] = []
        total_resources = 0
        total_notes = len(list(root.findall("note")))

        for idx, note_el in enumerate(root.findall("note"), 1):
            context.cancellation.check()
            context.progress.report_progress(
                10.0 + 60.0 * (idx / max(1, total_notes)),
                f"Processing note {idx}/{total_notes}...",
            )

            title, content, resources = self._parse_note(note_el)
            note_title = (title or "").strip() or f"Note {idx}"
            note_titles.append(note_title)

            # Normalize ENML to HTML
            html_content = self._normalize_enml_to_html(content)
            referenced_hashes = {
                match.lower() for match in re.findall(r"__DOCWEN_RES_([0-9a-fA-F]{32})__", html_content)
            }
            missing_hashes = sorted(referenced_hashes.difference(resources))
            if missing_hashes:
                raise ValueError(
                    f"ENEX note {idx} references missing or corrupt resources: " + ", ".join(missing_hashes)
                )
            # Convert HTML to Markdown
            md_body = md_convert(html_content or "")

            # Inject resource references (aux_artifacts appended in-place)
            res_map = self._write_resources(
                resources,
                context,
                idx,
                aux_artifacts,
                keep_images=keep_images,
                enable_ocr=enable_ocr,
                image_mode=image_mode,
                ocr_placement=ocr_placement,
                ocr_language=ocr_language,
                current_locale=current_locale,
                image_link_style=image_link_style,
            )
            for md5, replacement in res_map.items():
                token = f"__DOCWEN_RES_{md5}__"
                md_body = _replace_resource_token(md_body, token, replacement)

            total_resources += len(resources)

            md_note = f"# {note_title}\n\n{md_body.strip()}"
            notes.append(md_note)

        if not notes:
            notes.append("# (empty ENEX file)\n\n*No notes found in this ENEX file.*")

        from docwen_core.yaml_tools import generate_basic_yaml_frontmatter

        frontmatter_title = note_titles[0] if len(note_titles) == 1 else Path(input_path).stem
        yaml_frontmatter = generate_basic_yaml_frontmatter(
            frontmatter_title,
            yaml_key_labels=yaml_key_labels,
        )
        md_content = yaml_frontmatter + "\n\n---\n\n".join(notes).strip() + "\n"
        stats = {"notes": len(notes), "resources": total_resources}
        return md_content, stats

    @staticmethod
    def _parse_note(note_el: Any) -> tuple[str, str, dict[str, _EnexResource]]:
        """Parse a single ``<note>`` element.

        Returns (title, content, resources_dict).
        """
        title = _text(note_el.find("title")) or ""
        content = _text(note_el.find("content")) or ""

        resources: dict[str, _EnexResource] = {}
        for resource_index, res in enumerate(note_el.findall("resource"), 1):
            mime = _text(res.find("mime")) or "application/octet-stream"
            data_el = res.find("data")
            if data_el is None or (data_el.text or "").strip() == "":
                raise ValueError(f"ENEX resource {resource_index} is missing base64 data")
            encoding = (data_el.get("encoding") or "base64").strip().lower()
            if encoding != "base64":
                raise ValueError(f"ENEX resource {resource_index} uses unsupported encoding: {encoding or '(empty)'}")
            raw_b64 = re.sub(r"\s+", "", data_el.text or "")
            try:
                blob = base64.b64decode(raw_b64, validate=True)
            except (ValueError, binascii.Error) as exc:
                raise ValueError(f"ENEX resource {resource_index} contains invalid base64 data") from exc
            if not blob:
                raise ValueError(f"ENEX resource {resource_index} decodes to empty data")

            md5 = hashlib.md5(blob).hexdigest()
            attrs = res.find("resource-attributes")
            file_name = _text(attrs.find("file-name")) if attrs is not None else None

            resources[md5] = _EnexResource(md5=md5, mime=mime, data=blob, file_name=file_name)

        return title, content, resources

    @staticmethod
    def _normalize_enml_to_html(enml: str) -> str:
        """Normalize Evernote ENML to the HTML fragment sent to markdownify.

        ``<content>`` stores XML-wrapped ENML, while markdownify expects an
        HTML fragment.  Strip the XML envelope and replace ``<en-media>`` tags
        with ``<img>`` placeholder tokens that are resolved after conversion.
        """
        if not enml:
            return ""
        s = enml
        s = re.sub(r"^\s*<\?xml[^>]*\?>", "", s, flags=re.IGNORECASE)
        s = re.sub(r"^\s*<!DOCTYPE[^>]*>", "", s, flags=re.IGNORECASE)
        s = re.sub(r"^\s*<en-note\b[^>]*>", "", s, flags=re.IGNORECASE)
        s = re.sub(r"</en-note>\s*$", "", s, flags=re.IGNORECASE)
        s = re.sub(
            r'<en-media[^>]*\bhash="([0-9a-fA-F]{32})"[^>]*/?>',
            lambda m: f'<img src="__DOCWEN_RES_{m.group(1).lower()}__" />',
            s,
            flags=re.IGNORECASE,
        )
        return s.strip()

    @staticmethod
    def _write_resources(
        resources: dict[str, _EnexResource],
        context: ConverterContext,
        note_idx: int,
        aux_artifacts: list[Any],
        *,
        keep_images: bool,
        enable_ocr: bool,
        image_mode: str,
        ocr_placement: str,
        ocr_language: str,
        current_locale: str,
        image_link_style: str | None,
    ) -> dict[str, str]:
        """Write embedded ENEX resources to staging and populate aux_artifacts.

        Returns a dict mapping MD5 hashes to Markdown image links.
        Appends each resource's ArtifactManifest to aux_artifacts.
        Uses the shared ``MarkdownResourceWriter`` for naming, staging,
        artifact registration, and link formatting.
        """
        from docwen_plugin_markup.markdown_resources import (
            MarkdownResource,
            MarkdownResourceWriter,
        )

        out: dict[str, str] = {}
        if not keep_images and not enable_ocr:
            return dict.fromkeys(resources, "")

        resources_to_write: list[MarkdownResource] = []
        for r_idx, (md5, res) in enumerate(resources.items(), 1):
            ext = _ext_from_resource(res)
            filename = res.file_name or f"enex_note{note_idx}_res{r_idx}.{ext}"
            resources_to_write.append(
                MarkdownResource(
                    source_key=md5,
                    suggested_name=filename,
                    media_type=res.mime,
                    data=res.data,
                )
            )

        written = MarkdownResourceWriter().write_all(
            context,
            resources_to_write,
            image_link_style=image_link_style,
            image_mode=image_mode,
            enable_ocr=enable_ocr,
            ocr_placement=ocr_placement,
            ocr_language=ocr_language,
            current_locale=current_locale,
            source_format="enex",
            keep_resource_artifacts=keep_images,
        )
        for md5 in resources:
            item = written.get(md5.lower()) or written.get(md5)
            if item is None:
                continue
            aux_artifacts.extend(item.artifacts)
            out[md5] = item.markdown_link

        return out


def _text(el: Any) -> str | None:
    """Safely extract text content from an XML element."""
    if el is None:
        return None
    value = el.text
    if value is None:
        return None
    return str(value)


def _ext_from_resource(res: _EnexResource) -> str:
    """Determine a file extension from an ENEX resource."""
    if res.file_name:
        suffix = os.path.splitext(res.file_name)[1].lstrip(".").lower()
        if suffix:
            return suffix
    mime = (res.mime or "").lower()
    mime_map = {
        "image/jpeg": "jpeg",
        "image/jpg": "jpeg",
        "image/png": "png",
        "image/gif": "gif",
        "image/tiff": "tiff",
        "image/webp": "webp",
        "application/pdf": "pdf",
    }
    return mime_map.get(mime, "bin")


def _replace_resource_token(markdown: str, token: str, replacement: str) -> str:
    pattern = re.compile(r"!\[[^\]]*\]\(\s*" + re.escape(token) + r"\s*\)")
    markdown = pattern.sub(replacement, markdown)
    return markdown.replace(token, replacement)


def _make_aux_artifact(artifact_id: str, staging_path: str, filename: str, mime: str) -> Any:
    """Build an auxiliary ArtifactManifest for an extracted resource."""
    from docwen_core.models.artifact import ArtifactManifest

    return ArtifactManifest(
        artifact_id=artifact_id,
        kind="auxiliary",
        staging_path=staging_path,
        suggested_name=filename,
        media_type=mime,
        is_primary=False,
    )
