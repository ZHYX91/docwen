"""HTML / MHTML → Markdown converter.

Handles four source formats sharing the same conversion logic:
- ``html`` — standalone HTML files
- ``htm`` — HTML alias (same as html)
- ``mhtml`` — MIME HTML (multipart web archive)
- ``mht`` — MHTML alias (same as mhtml)

Uses ``markdownify`` to convert HTML body content to Markdown.
For MHTML, the multipart container is unpacked via stdlib ``email``
before the HTML payload is extracted and converted.

The converter:
- Only writes to staging via ``WorkspaceHandle``.
- Checks cancellation before expensive operations.
- Reports progress through ``ProgressSink``.
- Returns a ``ConversionResult`` with ``ArtifactManifest`` entries.
"""

from __future__ import annotations

import contextlib
import mimetypes
import os
import re
import time
import uuid
from email import policy
from email.parser import BytesParser
from html import unescape
from importlib.util import find_spec
from pathlib import Path
from typing import TYPE_CHECKING, Any
from urllib.parse import unquote, urljoin, urlparse

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


class HtmlToMarkdownConverter:
    """Convert HTML or MHTML to Markdown.

    Uses ``markdownify`` for the HTML→MD transformation.
    MHTML archives are unpacked to extract the HTML payload and
    embedded resources before conversion.
    """

    def convert(self, context: ConverterContext) -> Any:
        """Run the HTML/MHTML → Markdown conversion.

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
                    message=(
                        "HTML/MHTML→MD requires the 'markdownify' library. Install it with: pip install markdownify"
                    ),
                    diagnostic_code="HTML2MD-MISSING-DEP",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message="Missing required dependency: markdownify",
                        code="HTML2MD-MISSING-DEP",
                    ),
                ],
            )

        # 3. Determine source format from the request
        source = context.request.input_refs[0].format if context.request.input_refs else "html"

        # 4. Report start
        context.progress.report_progress(0.0, f"Starting {source.upper()} → Markdown conversion")
        context.logger.info(f"{source.upper()}→MD: reading {input_path}")

        # 5. Parse and convert
        context.cancellation.check()
        context.progress.report_progress(10.0, "Parsing input...")

        try:
            md_content, stats, aux_artifacts = self._convert(context, input_path, source)
        except Exception as exc:
            context.logger.error(f"{source.upper()}→MD conversion failed: {exc}")
            from docwen_core.models.result import ConversionErrorInfo

            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="HTML2MD-PARSE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Failed to convert {source.upper()}: {exc}",
                        code="HTML2MD-PARSE-ERROR",
                    ),
                ],
            )

        # 6. Write markdown to staging
        context.cancellation.check()
        context.progress.report_progress(80.0, "Writing Markdown to staging...")

        output_path = context.workspace.create_artifact_path("primary", ".md")
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md_content)
        except OSError as exc:
            context.logger.error(f"{source.upper()}→MD write failed: {exc}")
            from docwen_core.models.result import ConversionErrorInfo

            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=f"Failed to write output file: {exc}",
                    diagnostic_code="HTML2MD-WRITE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"File write error at {output_path}: {exc}",
                        code="HTML2MD-WRITE-ERROR",
                    ),
                ],
            )

        # 7. Build artifact manifest
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
                "source_format": source,
                "title": stats.get("title", ""),
                "image_count": stats.get("images", 0),
            },
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)

        # 8. Report completion
        context.progress.report_artifact_ready(artifact.artifact_id, suggested_name)
        context.progress.report_progress(100.0, "Conversion complete")
        context.logger.info(f"{source.upper()}→MD complete: {stats.get('title', 'untitled')}")

        duration_ms = (time.monotonic() - t_start) * 1000.0

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact, *aux_artifacts],
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=(f"Converted {source.upper()} to Markdown"),
                    code="HTML2MD-OK",
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

    # ── Internal conversion ───────────────────────────────────────────

    def _convert(
        self,
        context: ConverterContext,
        input_path: str,
        source: str,
    ) -> tuple[str, dict[str, Any], list[Any]]:
        """Convert HTML/MHTML file to Markdown.

        Returns (markdown_text, stats_dict, auxiliary_artifacts).
        """
        from markdownify import markdownify as md_convert

        # Handle MHTML: unpack multipart archive
        html_path = input_path
        resource_dir: str | None = None
        aux_artifacts: list[Any] = []
        prewritten_image_links: dict[str, str] = {}
        options = context.request.options
        keep_images = bool(options.get("to_md_keep_images", True))
        enable_ocr = bool(options.get("to_md_enable_ocr", False))
        from docwen_core.export_semantics import get_markdown_export_modes, resolve_markdown_request_policy

        request_policy = resolve_markdown_request_policy(context)
        export_semantics = request_policy.export

        export_modes = get_markdown_export_modes(
            "markup",
            extraction_mode=options.get("image_mode"),
            ocr_placement_mode=options.get("ocr_placement"),
            semantics=export_semantics,
        )
        image_mode = str(export_modes["image_extraction_mode"]).strip().lower()
        if image_mode not in {"file", "base64", "embed", "omit"}:
            image_mode = "file"
        ocr_placement = str(export_modes["ocr_placement_mode"]).strip().lower()
        if ocr_placement not in {"image_md", "main_md"}:
            ocr_placement = "image_md"
        ocr_language = str(options.get("ocr_language") or "auto")
        current_locale = str(options.get("locale") or "zh_CN")

        image_link_style = str(options.get("image_link_style") or export_semantics.image_link_style)

        if source in ("mhtml", "mht"):
            html_path, resource_dir, aux_artifacts, prewritten_image_links = self._parse_mhtml(
                input_path,
                context,
                keep_images=keep_images,
                enable_ocr=enable_ocr,
                image_mode=image_mode,
                ocr_placement=ocr_placement,
                ocr_language=ocr_language,
                current_locale=current_locale,
                image_link_style=image_link_style,
            )
        elif source in ("html", "htm"):
            resource_dir = _detect_resource_folder(input_path)

        # Read HTML content
        html_bytes = Path(html_path).read_bytes()

        # Try to decode HTML
        try:
            html_text = html_bytes.decode("utf-8", errors="replace")
        except Exception:
            html_text = html_bytes.decode("latin-1", errors="replace")

        # Extract title from HTML
        title = _extract_html_title(html_text)

        # Convert HTML to Markdown
        context.cancellation.check()
        context.progress.report_progress(40.0, "Converting HTML to Markdown...")

        html_for_markdown, image_count, image_artifacts, token_to_link = _prepare_html_image_links(
            html_text=html_text,
            html_path=html_path,
            resource_dir=resource_dir,
            context=context,
            keep_images=keep_images,
            enable_ocr=enable_ocr,
            image_mode=image_mode,
            ocr_placement=ocr_placement,
            ocr_language=ocr_language,
            current_locale=current_locale,
            image_link_style=image_link_style,
            prewritten_image_links=prewritten_image_links,
        )
        aux_artifacts.extend(image_artifacts)
        md_body = md_convert(html_for_markdown)
        for token, link in token_to_link.items():
            md_body = md_body.replace(token, link)

        # Count images before link-style rendering so wiki links are counted too.

        # Build final output. Expose the document title through
        # YAML frontmatter; the body itself should stay body-only.
        from docwen_core.yaml_tools import generate_basic_yaml_frontmatter

        yaml_frontmatter = generate_basic_yaml_frontmatter(
            title or Path(input_path).stem,
            yaml_key_labels=options.get("yaml_key_labels"),
        )
        final_md = f"{yaml_frontmatter}{md_body.strip()}\n"

        stats: dict[str, Any] = {
            "title": title,
            "images": image_count,
            "source_format": source,
            "resource_dir": resource_dir,
            "resource_count": len(aux_artifacts),
        }
        return final_md, stats, aux_artifacts

    # ── MHTML parsing ─────────────────────────────────────────────────

    @staticmethod
    def _parse_mhtml(
        mhtml_path: str,
        context: ConverterContext,
        *,
        keep_images: bool,
        enable_ocr: bool,
        image_mode: str,
        ocr_placement: str,
        ocr_language: str,
        current_locale: str,
        image_link_style: str,
    ) -> tuple[str, str, list[Any], dict[str, str]]:
        """Unpack an MHTML archive and extract the HTML payload.

        Returns (html_file_path, resource_directory_path, auxiliary_artifacts, suggested_name_to_markdown_link).
        """
        from docwen_plugin_markup.markdown_resources import (
            MarkdownResource,
            MarkdownResourceWriter,
            normalize_resource_key,
        )

        raw = Path(mhtml_path).read_bytes()
        msg = BytesParser(policy=policy.default).parsebytes(raw)

        html_content: str | None = None
        replacements: dict[str, str] = {}
        resource_refs: list[tuple[MarkdownResource, list[str]]] = []

        parts = list(msg.walk()) if msg.is_multipart() else [msg]
        for idx, part in enumerate(parts):
            # Check cancellation periodically during MIME part extraction
            if idx % 10 == 0:
                context.cancellation.check()
            if part.is_multipart():
                continue

            ctype = (part.get_content_type() or "").lower()
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            assert isinstance(payload, bytes)

            if ctype == "text/html":
                if html_content is None:
                    candidate = _decode_mhtml_html_payload(
                        payload,
                        mime_charset=part.get_content_charset(),
                    )
                    if candidate.strip():
                        html_content = candidate
                continue

            content_location = part.get("Content-Location")
            content_id = part.get("Content-ID")
            filename = part.get_filename()

            name = ""
            if content_location:
                try:
                    parsed = urlparse(content_location)
                    name = Path(parsed.path).name or ""
                except Exception:
                    name = Path(str(content_location)).name
            if not name and filename:
                name = Path(filename).name
            if not name:
                ext = ""
                if "/" in ctype:
                    subtype = ctype.split("/", 1)[1]
                    if subtype:
                        ext = f".{subtype.split(';', 1)[0].strip()}"
                name = f"part{len(replacements) + 1}{ext}"

            refs: list[str] = []
            if content_location:
                refs.append(str(content_location))
            if content_id:
                cid = str(content_id).strip()
                cid = cid.strip("<>").strip()
                if cid:
                    refs.append(f"cid:{cid}")

            source_key = refs[0] if refs else name or f"part{idx}"
            resource_refs.append(
                (
                    MarkdownResource(
                        source_key=source_key,
                        suggested_name=name,
                        media_type=ctype or "application/octet-stream",
                        data=payload,
                    ),
                    refs,
                )
            )

        if html_content is None:
            raise ValueError("MHTML archive does not contain a usable text/html body")

        # Build the extracted-resource directory only after the archive has
        # passed its required-content boundary. Invalid inputs must not create
        # a primary artifact or silently degrade to filename-only frontmatter.
        marker_path = context.workspace.create_artifact_path("auxiliary", ".tmp")
        staging_root = Path(marker_path).parent
        tmp_dir = staging_root / f"_mhtml_{uuid.uuid4().hex[:8]}"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        with contextlib.suppress(OSError):
            Path(marker_path).unlink(missing_ok=True)

        aux_artifacts: list[Any] = []
        prewritten_image_links: dict[str, str] = {}
        if keep_images or enable_ocr:
            image_sources = _extract_html_image_sources(html_content or "")
            normalized_image_sources = {normalize_resource_key(source) for source in image_sources}
            referenced_resources: list[tuple[MarkdownResource, list[str]]] = []
            for resource, refs in resource_refs:
                if not resource.media_type.startswith("image/"):
                    continue
                candidates = [resource.source_key, resource.suggested_name, *refs]
                if any(
                    candidate in image_sources or normalize_resource_key(candidate) in normalized_image_sources
                    for candidate in candidates
                ):
                    referenced_resources.append((resource, refs))

            resources = [resource for resource, _refs in referenced_resources]
            written_resources = MarkdownResourceWriter().write_all(
                context,
                resources,
                image_link_style=image_link_style,
                image_mode=image_mode,
                enable_ocr=enable_ocr,
                ocr_placement=ocr_placement,
                ocr_language=ocr_language,
                current_locale=current_locale,
                source_format="mhtml",
                keep_resource_artifacts=keep_images,
            )
            for resource, refs in referenced_resources:
                item = written_resources.get(normalize_resource_key(resource.source_key))
                if item is None:
                    continue
                aux_artifacts.extend(item.artifacts)
                prewritten_image_links[item.suggested_name] = item.markdown_link
                for ref in refs:
                    replacements[ref] = item.suggested_name

        # Apply replacements in HTML content
        for k, v in replacements.items():
            html_content = html_content.replace(k, v)

        html_output_path = tmp_dir / "content.html"
        html_output_path.write_text(html_content, encoding="utf-8")

        return str(html_output_path), str(tmp_dir), aux_artifacts, prewritten_image_links


def _prepare_html_image_links(
    *,
    html_text: str,
    html_path: str,
    resource_dir: str | None,
    context: ConverterContext,
    keep_images: bool,
    enable_ocr: bool,
    image_mode: str,
    ocr_placement: str,
    ocr_language: str,
    current_locale: str,
    image_link_style: str,
    prewritten_image_links: dict[str, str],
) -> tuple[str, int, list[Any], dict[str, str]]:
    """Replace HTML ``<img>`` elements with tokens rendered by shared link semantics."""
    try:
        import lxml.etree as etree
        import lxml.html as lxml_html
    except Exception:
        return html_text, len(re.findall(r"<img\b", html_text, re.IGNORECASE)), [], {}

    try:
        parser = lxml_html.HTMLParser(encoding="utf-8")
        doc = lxml_html.document_fromstring(html_text.encode("utf-8", errors="replace"), parser=parser)
    except Exception:
        return html_text, len(re.findall(r"<img\b", html_text, re.IGNORECASE)), [], {}

    from docwen_plugin_markup.markdown_resources import (
        MarkdownResource,
        MarkdownResourceWriter,
        normalize_resource_key,
    )

    base_href = _extract_base_href_from_doc(doc)
    token_to_link: dict[str, str] = {}
    resource_entries: list[tuple[str, MarkdownResource]] = []
    image_count = 0

    for image_count, img_el in enumerate(list(doc.iter("img")), 1):
        src = (img_el.get("src") or "").strip()
        if not src:
            continue

        token = f"DOCWENHTMLIMAGE{image_count}TOKEN"
        token_to_link[token] = ""

        if keep_images or enable_ocr:
            prewritten_link = prewritten_image_links.get(src)
            if prewritten_link is not None:
                token_to_link[token] = prewritten_link
            else:
                immediate_link, resource = _html_image_link_or_resource(
                    token=token,
                    src=src,
                    html_path=html_path,
                    base_href=base_href,
                    resource_dir=resource_dir,
                    image_link_style=image_link_style,
                )
                if resource is None:
                    token_to_link[token] = immediate_link if keep_images else ""
                else:
                    resource_entries.append((token, resource))

        _replace_node_with_text(etree, img_el, token)

    aux_artifacts: list[Any] = []
    if resource_entries:
        writer = MarkdownResourceWriter()
        written = writer.write_all(
            context,
            [resource for _token, resource in resource_entries],
            image_link_style=image_link_style,
            image_mode=image_mode,
            enable_ocr=enable_ocr,
            ocr_placement=ocr_placement,
            ocr_language=ocr_language,
            current_locale=current_locale,
            source_format="html",
            keep_resource_artifacts=keep_images,
        )
        for token, resource in resource_entries:
            item = written.get(normalize_resource_key(resource.source_key))
            if item is None:
                continue
            aux_artifacts.extend(item.artifacts)
            token_to_link[token] = item.markdown_link

    html_for_markdown = _html_body_or_document_for_markdown(lxml_html, doc)
    return html_for_markdown, image_count, aux_artifacts, token_to_link


def _html_body_or_document_for_markdown(lxml_html_module: Any, doc: Any) -> str:
    """Return HTML content that should become Markdown body text."""
    body = doc.find("body")
    target = body if body is not None else doc
    return lxml_html_module.tostring(target, encoding="unicode", method="html")


def _html_image_link_or_resource(
    *,
    token: str,
    src: str,
    html_path: str,
    base_href: str | None,
    resource_dir: str | None,
    image_link_style: str,
) -> tuple[str, Any | None]:
    """Return either a ready Markdown link or a resource to write through MarkdownResourceWriter."""
    from docwen_core.export_semantics import format_image_link
    from docwen_plugin_markup.markdown_resources import MarkdownResource

    if src.startswith("data:"):
        from docwen_core.links import is_data_uri_image, resolve_data_uri_image_to_temp_file

        if not is_data_uri_image(src):
            return "", None
        temp_file = resolve_data_uri_image_to_temp_file(src, temp_dir=None)
        if not temp_file:
            return "", None
        source_path = Path(temp_file)
        return "", MarkdownResource(
            source_key=token,
            suggested_name=source_path.name,
            media_type=_guess_media_type(source_path.name),
            data=source_path.read_bytes(),
        )

    if _is_remote_url(src):
        return format_image_link(src, src, style=image_link_style), None

    if base_href and _is_remote_url(base_href):
        try:
            joined = urljoin(base_href, src)
            if _is_remote_url(joined):
                return format_image_link(joined, joined, style=image_link_style), None
        except Exception:
            pass

    local_path = _resolve_local_path(src=src, html_path=html_path, base_href=base_href, resource_dir=resource_dir)
    if local_path is None or not local_path.exists() or not local_path.is_file():
        return format_image_link(src, src, style=image_link_style), None

    return "", MarkdownResource(
        source_key=token,
        suggested_name=local_path.name,
        media_type=_guess_media_type(local_path.name),
        data=local_path.read_bytes(),
    )


def _replace_node_with_text(etree_module: Any, node: Any, text: str) -> None:
    parent = node.getparent()
    if parent is None:
        node.text = text
        for key in list(node.attrib.keys()):
            node.attrib.pop(key, None)
        node.tag = "p"
        return
    placeholder = etree_module.Element("p")
    placeholder.text = text
    placeholder.tail = node.tail
    parent.replace(node, placeholder)


def _extract_base_href_from_doc(doc: Any) -> str | None:
    try:
        base_el = doc.find(".//base")
        if base_el is None:
            return None
        href = (base_el.get("href") or "").strip()
        return href or None
    except Exception:
        return None


def _detect_resource_folder(html_path: str) -> str | None:
    """Detect companion resource folder for an HTML file.

    Looks for ``<stem>_files`` or ``<stem>.files`` directories
    next to the HTML file.
    """
    p = Path(html_path)
    if not p.exists():
        return None
    stem = p.stem
    parent = p.parent
    candidates = [parent / f"{stem}_files", parent / f"{stem}.files"]
    for c in candidates:
        if c.is_dir():
            return str(c)
    return None


def _is_remote_url(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"}


def _resolve_local_path(*, src: str, html_path: str, base_href: str | None, resource_dir: str | None) -> Path | None:
    parsed = urlparse(src)
    if parsed.scheme == "file":
        try:
            return Path(unquote(parsed.path.lstrip("/")))
        except Exception:
            return None

    if parsed.scheme:
        return None

    raw_path = parsed.path or src
    raw_path = unquote(raw_path).replace("\\", "/")
    base_dir = Path(resource_dir) if resource_dir else Path(html_path).parent

    if raw_path.startswith("/") and base_href and base_href.startswith("file:"):
        try:
            base_parsed = urlparse(base_href)
            root = Path(unquote(base_parsed.path.lstrip("/"))).parent
            return root / raw_path.lstrip("/")
        except Exception:
            return None

    return (base_dir / raw_path).resolve()


def _guess_media_type(filename: str) -> str:
    media_type, _encoding = mimetypes.guess_type(filename)
    return media_type or "application/octet-stream"


def _extract_html_title(html_text: str) -> str:
    """Extract the title from an HTML document."""
    match = re.search(r"<title[^>]*>(.*?)</title>", html_text, re.IGNORECASE | re.DOTALL)
    if match:
        return unescape(match.group(1).strip())
    return ""


def _decode_mhtml_html_payload(payload: bytes, *, mime_charset: str | None) -> str:
    """Decode an MHTML HTML part using MIME and in-document declarations."""
    declared_match = re.search(
        rb"charset\s*=\s*['\"]?\s*([a-zA-Z0-9._:-]+)",
        payload[:16384],
        re.IGNORECASE,
    )
    declared_charset = declared_match.group(1).decode("ascii") if declared_match else None
    candidates = [mime_charset, declared_charset, "utf-8", "windows-1252"]
    attempted: set[str] = set()
    for candidate in candidates:
        if not candidate:
            continue
        normalized = candidate.strip().lower()
        if not normalized or normalized in attempted:
            continue
        attempted.add(normalized)
        try:
            return payload.decode(normalized)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode("utf-8", errors="replace")


def _extract_html_image_sources(html_text: str) -> set[str]:
    """Return only sources consumed by the converter's ``img`` projection."""
    try:
        from lxml import html as lxml_html

        document = lxml_html.fromstring(html_text)
        return {source for element in document.iter("img") if (source := str(element.get("src") or "").strip())}
    except Exception:
        return {
            match.group(2).strip()
            for match in re.finditer(
                r"<img\b[^>]*\bsrc\s*=\s*(['\"])(.*?)\1",
                html_text,
                re.IGNORECASE | re.DOTALL,
            )
            if match.group(2).strip()
        }
