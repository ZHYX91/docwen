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
from html import unescape
from importlib.util import find_spec
from pathlib import Path, PurePosixPath
from typing import TYPE_CHECKING, Any
from urllib.parse import unquote, urljoin, urlparse
from urllib.request import url2pathname

from docwen_core.text.placeholders import placeholder_namespace, restore_placeholders
from docwen_plugin_markup.web_archive.mime_structure import parse_mhtml_structure
from docwen_plugin_markup.web_archive.resource_identity import (
    content_id_identity,
    mhtml_resource_identity,
    resolve_mhtml_uri,
)

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

        missing = stats.get("missing_local_images", {})
        extra_diagnostics = []
        if missing.get("count", 0):
            positions = ", ".join(str(index) for index in missing["positions"])
            retention = (
                "Original links were retained."
                if missing.get("links_retained", True)
                else "No image links were emitted for these references."
            )
            if missing.get("ocr_requested", False):
                retention += " OCR could not run for these references."
            extra_diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    code="HTML-LOCAL-IMAGE-UNAVAILABLE",
                    message=(
                        f"{missing['count']} local image reference(s) could not be packaged; "
                        f"image positions: {positions}. {retention}"
                    ),
                )
            )
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
                *extra_diagnostics,
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
        mhtml_base_uri: str | None = None
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
            html_path, resource_dir, aux_artifacts, prewritten_image_links, mhtml_base_uri = self._parse_mhtml(
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

        # Read HTML content. MHTML extraction already writes normalized UTF-8;
        # ordinary HTML keeps its admitted encoding and in-document charset.
        html_bytes = Path(html_path).read_bytes()
        source_ref = context.request.input_refs[0] if context.request.input_refs else None
        admitted_encoding = (
            str(source_ref.encoding or "").strip() if source_ref is not None and source in {"html", "htm"} else "utf-8"
        )
        html_text = _decode_html_payload(html_bytes, admitted_encoding=admitted_encoding)

        # Extract title from HTML
        title = _extract_html_title(html_text)

        # Convert HTML to Markdown
        context.cancellation.check()
        context.progress.report_progress(40.0, "Converting HTML to Markdown...")

        html_for_markdown, image_count, image_artifacts, token_to_link, missing_local = _prepare_html_image_links(
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
            mhtml_base_uri=mhtml_base_uri,
        )
        aux_artifacts.extend(image_artifacts)
        md_body = md_convert(html_for_markdown)
        md_body = restore_placeholders(md_body, token_to_link)

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
            "missing_local_images": missing_local,
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
    ) -> tuple[str, str, list[Any], dict[str, str], str]:
        """Unpack an MHTML archive and extract the HTML payload.

        Returns the HTML path, resource directory, artifacts, exact-src links and selected HTML URI base.
        """
        from docwen_plugin_markup.markdown_resources import (
            MarkdownResource,
            MarkdownResourceWriter,
        )

        root_part, parts = parse_mhtml_structure(Path(mhtml_path).read_bytes())
        html_payload = root_part.message.get_payload(decode=True)
        if not isinstance(html_payload, bytes):
            raise ValueError("MHTML selected root does not contain a usable text/html body")
        html_content = _decode_mhtml_html_payload(html_payload, mime_charset=root_part.message.get_content_charset())
        if not html_content.strip():
            raise ValueError("MHTML selected root does not contain a usable text/html body")
        document_uri = root_part.base_uri
        resource_refs: list[tuple[MarkdownResource, list[str]]] = []
        alias_keys: dict[str, set[str]] = {}
        resources_by_key: dict[str, MarkdownResource] = {}
        needs_image_data = (keep_images and image_mode != "omit") or enable_ocr
        for idx, entry in enumerate(parts if needs_image_data else []):
            if idx % 10 == 0:
                context.cancellation.check()
            part = entry.message
            ctype = part.get_content_type().lower()
            if part.is_multipart() or not ctype.startswith("image/"):
                continue
            payload = part.get_payload(decode=True)
            if not isinstance(payload, bytes):
                continue
            filename = part.get_filename()
            name = _mhtml_basename(entry.location) if entry.location else ""
            if not name and filename:
                name = Path(filename).name
            if not name:
                name = f"part{idx}.{ctype.split('/', 1)[-1]}"
            refs = [mhtml_resource_identity(entry.location)] if entry.location else []
            content_id = str(part.get("Content-ID") or "").strip().strip("<>")
            if content_id:
                refs.append(content_id_identity(content_id))
            resource = MarkdownResource(
                source_key=f"mhtml-part-{idx}", suggested_name=name, media_type=ctype, data=payload
            )
            resource_refs.append((resource, refs))
            resources_by_key[resource.source_key] = resource
            # Basename compatibility is separate from URI/CID identity and is
            # available only for genuinely named resources, never generated names.
            alias = _mhtml_basename(entry.location) if entry.location else (Path(filename).name if filename else "")
            if alias:
                alias_keys.setdefault(alias, set()).add(resource.source_key)

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
        if needs_image_data:
            import lxml.html as lxml_html

            doc = lxml_html.document_fromstring(
                html_content.encode("utf-8"), parser=lxml_html.HTMLParser(encoding="utf-8")
            )
            base_uri = resolve_mhtml_uri(document_uri, _extract_base_href_from_doc(doc) or "")
            image_sources = _extract_html_image_sources(html_content)
            by_identity: dict[str, MarkdownResource] = {}
            ambiguous: set[str] = set()
            for resource, refs in resource_refs:
                if not resource.media_type.startswith("image/"):
                    continue
                for key in refs:
                    previous = by_identity.get(key)
                    if previous is not None and previous.data != resource.data:
                        ambiguous.add(key)
                    by_identity[key] = resource
            source_resources: dict[str, MarkdownResource] = {}
            for source_ref in image_sources:
                key = mhtml_resource_identity(source_ref, base_uri=base_uri)
                if key in ambiguous:
                    raise ValueError("MHTML archive has conflicting referenced image resource identities")
                resource = by_identity.get(key)
                if resource is None:
                    parsed_source = urlparse(source_ref)
                    # A bare relative filename may use one unique alias. A
                    # missing absolute URI, CID, path or query cannot discard
                    # its identity by falling back to a similarly named image.
                    if (
                        not parsed_source.scheme
                        and not parsed_source.netloc
                        and "?" not in source_ref
                        and "#" not in source_ref
                        and "/" not in parsed_source.path
                        and "\\" not in parsed_source.path
                    ):
                        candidates = alias_keys.get(_mhtml_basename(source_ref), set())
                        if len(candidates) > 1:
                            raise ValueError("MHTML referenced image basename is ambiguous")
                        if len(candidates) == 1:
                            selected_key = next(iter(candidates))
                            resource = resources_by_key[selected_key]
                if resource is not None:
                    source_resources[source_ref] = resource
            selected = {resource.source_key for resource in source_resources.values()}
            referenced_resources = [
                (resource, refs) for resource, refs in resource_refs if resource.source_key in selected
            ]

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
            for resource, _refs in referenced_resources:
                item = written_resources.get(resource.source_key)
                if item is not None:
                    aux_artifacts.extend(item.artifacts)
            for source_ref, resource in source_resources.items():
                item = written_resources.get(resource.source_key)
                if item is not None:
                    prewritten_image_links[source_ref] = item.markdown_link

        # Keep authored body/code text intact. The image pass binds only img src.
        html_output_path = tmp_dir / "content.html"
        html_output_path.write_text(html_content, encoding="utf-8")

        return str(html_output_path), str(tmp_dir), aux_artifacts, prewritten_image_links, document_uri


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
    mhtml_base_uri: str | None = None,
) -> tuple[str, int, list[Any], dict[str, str], dict[str, Any]]:
    """Replace HTML ``<img>`` elements with tokens rendered by shared link semantics."""
    try:
        import lxml.etree as etree
        import lxml.html as lxml_html
    except Exception:
        return html_text, len(re.findall(r"<img\b", html_text, re.IGNORECASE)), [], {}, {"count": 0, "positions": []}

    try:
        parser = lxml_html.HTMLParser(encoding="utf-8")
        doc = lxml_html.document_fromstring(html_text.encode("utf-8", errors="replace"), parser=parser)
    except Exception:
        return html_text, len(re.findall(r"<img\b", html_text, re.IGNORECASE)), [], {}, {"count": 0, "positions": []}

    from docwen_plugin_markup.markdown_resources import (
        MarkdownResource,
        MarkdownResourceWriter,
        normalize_resource_key,
    )

    base_href = _extract_base_href_from_doc(doc)
    if mhtml_base_uri is not None:
        base_href = resolve_mhtml_uri(mhtml_base_uri, base_href or "")
    token_to_link: dict[str, str] = {}
    resource_entries: list[tuple[str, MarkdownResource]] = []
    image_count = 0
    image_links_enabled = keep_images and image_mode != "omit"
    missing_local: dict[str, Any] = {
        "count": 0,
        "positions": [],
        "links_retained": image_links_enabled,
        "ocr_requested": enable_ocr,
    }
    authored = (
        html_text + etree.tostring(doc, encoding="unicode", method="text") + etree.tostring(doc, encoding="unicode")
    )
    namespace = placeholder_namespace(authored, prefix="DOCWENHTMLIMAGE")

    for image_count, img_el in enumerate(list(doc.iter("img")), 1):
        src = (img_el.get("src") or "").strip()
        if not src:
            continue

        token = f"{namespace}{image_count}TOKEN"
        token_to_link[token] = ""

        if image_links_enabled or enable_ocr:
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
                    token_to_link[token] = immediate_link if keep_images and image_mode != "omit" else ""
                    if _is_local_image_reference(src=src, html_path=html_path, base_href=base_href):
                        missing_local["count"] += 1
                        if len(missing_local["positions"]) < 3:
                            missing_local["positions"].append(image_count)
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
    return html_for_markdown, image_count, aux_artifacts, token_to_link, missing_local


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
        from docwen_core.links import decode_data_uri_image, is_data_uri_image

        if not is_data_uri_image(src):
            return "", None
        decoded = decode_data_uri_image(src)
        if decoded is None:
            return "", None
        payload, suffix = decoded
        suggested_name = f"data-uri{suffix}"
        return "", MarkdownResource(
            source_key=token,
            suggested_name=suggested_name,
            media_type=_guess_media_type(suggested_name),
            data=payload,
        )

    try:
        target_uri = _html_image_uri(src=src, html_path=html_path, base_href=base_href)
        if _is_remote_url(target_uri):
            return format_image_link(target_uri, target_uri, style=image_link_style), None
        local_path = _resolve_local_path(src=src, html_path=html_path, base_href=base_href, resource_dir=resource_dir)
        if local_path is None or not local_path.is_file():
            return format_image_link(src, src, style=image_link_style), None
        data = local_path.read_bytes()
    except (OSError, ValueError, RuntimeError):
        return format_image_link(src, src, style=image_link_style), None

    return "", MarkdownResource(
        source_key=token,
        suggested_name=local_path.name,
        media_type=_guess_media_type(local_path.name),
        data=data,
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
        base_el = next((element for element in doc.iter("base") if element.get("href") is not None), None)
        if base_el is None:
            return None
        href = (base_el.get("href") or "").strip()
        return href
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


def _mhtml_basename(reference: str) -> str:
    component = urlparse(reference).path.rsplit("/", 1)[-1]
    decoded = unquote(component)
    return decoded if decoded and "/" not in decoded and "\\" not in decoded else ""


def _is_local_image_reference(*, src: str, html_path: str, base_href: str | None) -> bool:
    if src.lower().startswith("data:"):
        return False
    try:
        return urlparse(_html_image_uri(src=src, html_path=html_path, base_href=base_href)).scheme == "file"
    except (OSError, ValueError, RuntimeError):
        return src.lower().startswith("file:") or not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", src)


def _html_image_uri(*, src: str, html_path: str, base_href: str | None) -> str:
    def join(base: str, reference: str) -> str:
        if urlparse(base).scheme == "file":
            reference = reference.replace("\\", "/")
            match = re.match(r"^/([A-Za-z]:)(?:/|$)", urlparse(base).path)
            if (
                match
                and reference.startswith("/")
                and not reference.startswith("//")
                and not re.match(r"^/[A-Za-z]:(?:/|$)", reference)
            ):
                reference = "/" + match.group(1) + reference
        return urljoin(base, reference)

    document_uri = Path(html_path).resolve().as_uri()
    effective_base = join(document_uri, base_href) if base_href is not None else document_uri
    return join(effective_base, src)


def _file_uri_path(uri: str) -> Path | None:
    parsed = urlparse(uri)
    authority = parsed.netloc
    if authority and authority.lower() != "localhost":
        # Network shares are Windows paths; never reinterpret one as a local
        # POSIX path or fetch an archive's remote resource.
        return Path(url2pathname(f"//{authority}{parsed.path}")) if os.name == "nt" else None
    if os.name != "nt" and re.match(r"^/[A-Za-z]:", parsed.path):
        return None
    return Path(url2pathname(parsed.path))


def _resolve_local_path(*, src: str, html_path: str, base_href: str | None, resource_dir: str | None) -> Path | None:
    target = _html_image_uri(src=src, html_path=html_path, base_href=base_href)
    if urlparse(target).scheme != "file":
        return None
    document_relative = _file_uri_path(target)
    if document_relative is None:
        return None
    parsed = urlparse(src)
    # An explicit base or file URI fixes resource identity; companion fallback
    # is only for ordinary relative references without any base declaration.
    if base_href is not None or parsed.scheme or document_relative.exists() or resource_dir is None:
        return document_relative
    raw_path = unquote(parsed.path or src).replace("\\", "/")

    resource_root = Path(resource_dir)
    normalized_parts = PurePosixPath(raw_path).parts
    if normalized_parts and normalized_parts[0].casefold() == resource_root.name.casefold():
        companion_relative = resource_root.parent.joinpath(*normalized_parts).resolve()
    else:
        companion_relative = resource_root.joinpath(*normalized_parts).resolve()
    return companion_relative


def _guess_media_type(filename: str) -> str:
    media_type, _encoding = mimetypes.guess_type(filename)
    return media_type or "application/octet-stream"


def _decode_html_payload(payload: bytes, *, admitted_encoding: str | None = None) -> str:
    """Decode ordinary HTML using admission plus in-document charset evidence."""

    declared_match = re.search(
        rb"(?:charset\s*=\s*[\"']?\s*)([a-zA-Z0-9._:-]+)",
        payload[:16384],
        re.IGNORECASE,
    )
    declared_charset = declared_match.group(1).decode("ascii") if declared_match else None

    bom_candidates: list[str] = []
    if payload.startswith(b"\xef\xbb\xbf"):
        bom_candidates.append("utf-8-sig")
    elif payload.startswith((b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff")):
        bom_candidates.append("utf-32")
    elif payload.startswith((b"\xff\xfe", b"\xfe\xff")):
        bom_candidates.append("utf-16")

    candidates = [
        *bom_candidates,
        admitted_encoding,
        declared_charset,
        "utf-8",
        "windows-1252",
        "latin-1",
    ]
    attempted: set[str] = set()
    for candidate in candidates:
        if not candidate:
            continue
        normalized = candidate.strip().lower()
        if not normalized or normalized in attempted:
            continue
        attempted.add(normalized)
        try:
            return payload.decode(candidate)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode("utf-8", errors="replace")


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
