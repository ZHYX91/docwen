"""EPUB → Markdown converter using ebooklib + beautifulsoup4."""

from __future__ import annotations

import time
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.artifact import (
    ARTIFACT_KIND_PRIMARY,
    ArtifactManifest,
)
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.paths import input_stem
from docwen_plugin_markup._common import file_size, new_artifact_id
from docwen_plugin_markup.markdown_resources import (
    MarkdownResource,
    MarkdownResourceWriter,
    normalize_resource_key,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


class EpubToMarkdownConverter:
    """Converts EPUB files to Markdown text.

    Uses ebooklib to parse the EPUB container and beautifulsoup4 to
    extract text content from HTML/XHTML chapters. Images can optionally
    be extracted and saved alongside the Markdown output.
    """

    def convert(self, context: ConverterContext) -> ConversionResult:
        task_id = context.request.request_id
        start_time = time.monotonic()
        input_file = context.workspace.input_path
        stem = input_stem(input_file)
        input_bytes = file_size(input_file)

        options = context.request.options or {}
        keep_images = bool(options.get("to_md_keep_images", True))
        enable_ocr = bool(options.get("to_md_enable_ocr", False))
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

        # ── Cancellation check ──────────────────────────────────────
        try:
            context.cancellation.check()
        except Exception:
            return self._cancelled(task_id)

        # ── Import dependencies ─────────────────────────────────────
        try:
            import ebooklib
            from ebooklib import epub
        except ImportError:
            return self._dependency_missing(
                task_id,
                "ebooklib is required for EPUB parsing. Install with: pip install ebooklib beautifulsoup4",
            )

        try:
            from bs4 import BeautifulSoup
        except ImportError:
            return self._dependency_missing(
                task_id,
                "beautifulsoup4 is required for HTML parsing. Install with: pip install ebooklib beautifulsoup4",
            )

        # ── Parse EPUB ──────────────────────────────────────────────
        try:
            book = epub.read_epub(input_file)
        except Exception as exc:
            return self._conversion_error(
                task_id,
                "EPUB2MD-PARSE-ERROR",
                f"Failed to parse EPUB file: {exc}",
            )

        images_written = 0
        degraded_diagnostics: list[ConversionDiagnostic] = []

        try:
            context.cancellation.check()
        except Exception:
            return self._cancelled(task_id)

        # ── Extract images if requested ─────────────────────────────
        image_links: dict[str, str] = {}
        image_artifacts: list[ArtifactManifest] = []
        image_items = list(book.get_items_of_type(ebooklib.ITEM_IMAGE))
        resources: list[MarkdownResource] = []
        for item in image_items:
            name = item.get_name() or f"image_{len(resources):03d}.jpg"
            content = item.get_content()
            if not content:
                continue
            resources.append(
                MarkdownResource(
                    source_key=name,
                    suggested_name=Path(name).name or f"image_{len(resources):03d}.jpg",
                    media_type=getattr(item, "media_type", "") or "application/octet-stream",
                    data=content,
                )
            )

        if keep_images or enable_ocr:
            written = MarkdownResourceWriter().write_all(
                context,
                resources,
                image_link_style=str(options.get("image_link_style") or "").strip() or None,
                image_mode=image_mode,
                enable_ocr=enable_ocr,
                ocr_placement=ocr_placement,
                ocr_language=ocr_language,
                current_locale=current_locale,
                source_format="epub",
                keep_resource_artifacts=keep_images,
            )
            seen_artifact_ids: set[str] = set()
            for key, item in written.items():
                image_links[key] = item.markdown_link
                for artifact in item.artifacts:
                    if artifact.artifact_id not in seen_artifact_ids:
                        image_artifacts.append(artifact)
                        seen_artifact_ids.add(artifact.artifact_id)
            for resource in resources:
                normalized = normalize_resource_key(resource.source_key)
                item = written.get(normalized)
                if item is None:
                    continue
                basename_key = normalize_resource_key(Path(resource.source_key).name)
                image_links.setdefault(basename_key, item.markdown_link)
            images_written = len({artifact.artifact_id for artifact in image_artifacts if artifact.kind == "image"})
        else:
            for resource in resources:
                image_links[normalize_resource_key(resource.source_key)] = ""
                image_links[normalize_resource_key(Path(resource.source_key).name)] = ""

        # ── Extract text content ────────────────────────────────────
        markdown_parts: list[str] = []

        # Title
        title = self._get_metadata(book, "title", degraded_diagnostics) or stem
        markdown_parts.append(f"# {title}\n")

        # Author
        author = self._get_metadata(book, "creator", degraded_diagnostics)
        if author:
            markdown_parts.append(f"*Author: {author}*\n")

        # Table of contents
        toc = book.toc
        if toc:
            markdown_parts.append("## Table of Contents\n")
            for toc_item in toc:
                if isinstance(toc_item, tuple):
                    # (section, children)
                    section = toc_item[0]
                    if hasattr(section, "title"):
                        markdown_parts.append(f"- {section.title}")
                    children = toc_item[1] if len(toc_item) > 1 else []
                    if children:
                        for child in children:
                            if hasattr(child, "title"):
                                markdown_parts.append(f"  - {child.title}")
                elif hasattr(toc_item, "title"):
                    markdown_parts.append(f"- {toc_item.title}")
            markdown_parts.append("")

        # Extract document items in spine order
        spine_items = []
        spine_lookup_failures = 0
        for item_id, _linear in book.spine:
            try:
                item = book.get_item_with_id(item_id)
                if item:
                    spine_items.append(item)
            except Exception as exc:
                spine_lookup_failures += 1
                degraded_diagnostics.append(
                    ConversionDiagnostic(
                        level="warning",
                        message=(
                            f"Could not resolve EPUB spine item '{item_id}' after "
                            f"{type(exc).__name__}; available document items will be used as fallback."
                        ),
                        code="EPUB2MD-SPINE-ITEM-FALLBACK",
                        location=str(item_id),
                    )
                )
                continue

        # Also include HTML items not in spine (fallback)
        document_items = list(book.get_items_of_type(ebooklib.ITEM_DOCUMENT))

        # Use spine order when complete.  If any lookup failed, retain the
        # resolved order and append remaining document items so a malformed
        # spine cannot silently drop readable chapters.
        if spine_items and spine_lookup_failures:
            items_to_process = list(spine_items)
            items_to_process.extend(
                item for item in document_items if all(item is not existing for existing in items_to_process)
            )
        else:
            items_to_process = spine_items if spine_items else document_items

        for item in items_to_process:
            try:
                context.cancellation.check()
            except Exception:
                return self._cancelled(task_id)

            item_location = str(getattr(item, "id", "") or "unknown EPUB item")
            try:
                if _is_navigation_item(item):
                    continue
                content = item.get_content()
                if not content:
                    continue
                html_content = content.decode("utf-8", errors="replace")
                soup = BeautifulSoup(html_content, "html.parser")

                # Remove script and style elements
                for tag in soup(["script", "style", "nav"]):
                    tag.decompose()

                # Get text
                text_root = soup.body if soup.body is not None else soup
                text = self._html_to_markdown(text_root, image_links)

                if text.strip():
                    # Add section heading from item name if available
                    item_name = item.get_name() or ""
                    if item_name:
                        section_title = Path(item_name).stem
                        # Only add if the title is meaningful
                        if (
                            section_title
                            and not section_title.startswith("index")
                            and len(section_title) > 2
                            and section_title.lower() != "cover"
                        ):
                            markdown_parts.append(f"## {section_title}\n")

                    markdown_parts.append(text)
                    markdown_parts.append("")

            except Exception as exc:
                degraded_diagnostics.append(
                    ConversionDiagnostic(
                        level="warning",
                        message=(
                            f"Skipped one unreadable EPUB chapter after {type(exc).__name__}; "
                            "the remaining chapters were preserved."
                        ),
                        code="EPUB2MD-CHAPTER-SKIPPED",
                        location=item_location,
                    )
                )
                continue

        # ── Write Markdown output ───────────────────────────────────
        from docwen_core.yaml_tools import generate_basic_yaml_frontmatter

        yaml_frontmatter = generate_basic_yaml_frontmatter(
            title,
            yaml_key_labels=options.get("yaml_key_labels"),
        )
        markdown_content = yaml_frontmatter + "\n".join(markdown_parts).strip() + "\n"

        staging_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".md")

        try:
            Path(staging_path).write_text(markdown_content, encoding="utf-8")
        except Exception as exc:
            return self._conversion_error(
                task_id,
                "EPUB2MD-WRITE-ERROR",
                f"Failed to write Markdown output: {exc}",
            )

        output_bytes = file_size(staging_path)

        artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=staging_path,
            suggested_name=f"{stem}.md",
            media_type="text/markdown",
            metadata={
                "title": title,
                "author": author or "",
                "image_count": images_written,
                "degradation_count": len(degraded_diagnostics),
            },
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)

        elapsed_ms = (time.monotonic() - start_time) * 1000.0
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact, *image_artifacts],
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=f"Converted EPUB to Markdown: '{title}'",
                    code="EPUB2MD-CONVERTED",
                ),
                *degraded_diagnostics,
            ],
            metrics=ConversionMetrics(
                duration_ms=elapsed_ms,
                input_bytes=input_bytes,
                output_bytes=output_bytes,
                extra={
                    "image_count": images_written,
                    "degradation_count": len(degraded_diagnostics),
                    "spine_lookup_failure_count": spine_lookup_failures,
                    "chapter_skip_count": sum(
                        diagnostic.code == "EPUB2MD-CHAPTER-SKIPPED" for diagnostic in degraded_diagnostics
                    ),
                },
            ),
        )

    # ── Helpers ──────────────────────────────────────────────────────

    @staticmethod
    def _get_metadata(
        book,
        key: str,
        diagnostics: list[ConversionDiagnostic],
    ) -> str:
        """Extract metadata from an ebooklib EPUB book."""
        try:
            items = book.get_metadata("DC", key)
            if items:
                # items are (value, attrib_dict) tuples
                return str(items[0][0])
        except Exception as exc:
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        f"Could not read EPUB metadata '{key}' after {type(exc).__name__}; "
                        "used the documented fallback."
                    ),
                    code="EPUB2MD-METADATA-FALLBACK",
                    location=f"metadata:{key}",
                )
            )
        return ""

    @staticmethod
    def _html_to_markdown(soup, image_links: dict[str, str] | None = None) -> str:
        """Convert BeautifulSoup HTML tree to Markdown text."""
        result: list[str] = []
        _html_to_md_recursive(soup, result, image_links or {})
        # Collapse multiple blank lines
        text = "\n".join(result)
        while "\n\n\n" in text:
            text = text.replace("\n\n\n", "\n\n")
        return text

    @staticmethod
    def _cancelled(task_id: str) -> ConversionResult:
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="cancelled",
                message="Conversion was cancelled by the user.",
                diagnostic_code="EPUB2MD-CANCELLED",
            ),
        )

    @staticmethod
    def _dependency_missing(task_id: str, message: str) -> ConversionResult:
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="dependency_missing",
                message=message,
                diagnostic_code="EPUB2MD-DEPENDENCY-MISSING",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=message,
                    code="EPUB2MD-DEPENDENCY-MISSING",
                )
            ],
        )

    @staticmethod
    def _conversion_error(task_id: str, code: str, message: str) -> ConversionResult:
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message=message,
                diagnostic_code=code,
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=message,
                    code=code,
                )
            ],
        )


# ── Recursive HTML-to-Markdown conversion ────────────────────────────

_BLOCK_TAGS = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote", "pre", "br", "hr", "tr"}

_HEADING_TAGS = {"h1": "#", "h2": "##", "h3": "###", "h4": "####", "h5": "#####", "h6": "######"}

_INLINE_TAGS_PREFIX = {"b": "**", "strong": "**", "i": "*", "em": "*"}

_INLINE_TAGS_NO_PREFIX = {"a", "span", "code", "sub", "sup", "small", "u", "s", "del", "ins", "mark"}


def _html_to_md_recursive(element, result: list[str], image_links: dict[str, str]) -> None:
    """Recursively convert an HTML element tree to Markdown lines."""
    from bs4 import NavigableString, Tag  # pyright: ignore[reportPrivateImportUsage]

    if isinstance(element, NavigableString):
        text = str(element)
        if text.strip():
            result.append(text.strip())
        return

    if not isinstance(element, Tag):
        return

    tag_name = element.name.lower() if element.name else ""

    if tag_name in _HEADING_TAGS:
        prefix = _HEADING_TAGS[tag_name] + " "
        text_parts: list[str] = []
        for child in element.children:
            _collect_text(child, text_parts, image_links)
        heading_text = "".join(text_parts).strip()
        if heading_text:
            result.append(prefix + heading_text)
            result.append("")

    elif tag_name == "p":
        text_parts = []
        for child in element.children:
            _collect_text(child, text_parts, image_links)
        para_text = "".join(text_parts).strip()
        if para_text:
            result.append(para_text)
            result.append("")

    elif tag_name == "li":
        text_parts = []
        for child in element.children:
            _collect_text(child, text_parts, image_links)
        li_text = "".join(text_parts).strip()
        if li_text:
            result.append("- " + li_text)

    elif tag_name in ("ul", "ol"):
        for child in element.children:
            _html_to_md_recursive(child, result, image_links)
        result.append("")

    elif tag_name == "blockquote":
        text_parts = []
        for child in element.children:
            _collect_text(child, text_parts, image_links)
        bq_text = "".join(text_parts).strip()
        if bq_text:
            for line in bq_text.split("\n"):
                result.append("> " + line)
            result.append("")

    elif tag_name == "pre":
        code = element.get_text()
        result.append("```")
        result.append(code.strip())
        result.append("```")
        result.append("")

    elif tag_name == "br":
        result.append("")

    elif tag_name == "hr":
        result.append("---")
        result.append("")

    elif tag_name == "table":
        rows = element.find_all("tr")
        for row_index, row in enumerate(rows):
            cells = row.find_all(["td", "th"])
            if not cells:
                continue
            cell_texts = [c.get_text().strip() for c in cells]
            result.append("| " + " | ".join(cell_texts) + " |")
            if row_index == 0 and any(getattr(c, "name", "").lower() == "th" for c in cells):
                result.append("| " + " | ".join("---" for _ in cells) + " |")
        result.append("")

    elif tag_name == "img":
        alt = str(element.get("alt") or "")
        src = str(element.get("src") or "")
        if src:
            key = normalize_resource_key(src)
            link = image_links.get(key)
            if link is not None:
                link = _replace_markdown_image_alt(link, alt)
                if link:
                    result.append(link)
            else:
                result.append(f"![{alt}]({src})")

    elif tag_name == "a":
        text_parts = []
        for child in element.children:
            _collect_text(child, text_parts, image_links)
        link_text = "".join(text_parts).strip()
        href = element.get("href", "")
        if link_text and href:
            result.append(f"[{link_text}]({href})")
        elif link_text:
            result.append(link_text)

    else:
        # Default: recurse into children
        for child in element.children:
            _html_to_md_recursive(child, result, image_links)


def _collect_text(element, result: list[str], image_links: dict[str, str]) -> None:
    """Collect visible text from an element tree, preserving inline formatting."""
    from bs4 import NavigableString, Tag  # pyright: ignore[reportPrivateImportUsage]

    if isinstance(element, NavigableString):
        result.append(str(element))
    elif isinstance(element, Tag):
        tag_name = element.name.lower() if element.name else ""

        if tag_name in _INLINE_TAGS_PREFIX:
            prefix = _INLINE_TAGS_PREFIX[tag_name]
            result.append(prefix)
            for child in element.children:
                _collect_text(child, result, image_links)
            result.append(prefix)
        elif tag_name == "code":
            result.append("`")
            for child in element.children:
                _collect_text(child, result, image_links)
            result.append("`")
        elif tag_name == "img":
            alt = str(element.get("alt") or "")
            src = str(element.get("src") or "")
            if src:
                key = normalize_resource_key(src)
                link = image_links.get(key)
                if link is not None:
                    link = _replace_markdown_image_alt(link, alt)
                    if link:
                        result.append(link)
                else:
                    result.append(f"![{alt}]({src})")
        elif tag_name == "a":
            href = element.get("href", "")
            if href:
                result.append("[")
            for child in element.children:
                _collect_text(child, result, image_links)
            if href:
                result.append(f"]({href})")
        elif tag_name == "br":
            result.append("\n")
        else:
            for child in element.children:
                _collect_text(child, result, image_links)


def _is_navigation_item(item) -> bool:
    """Return true for EPUB navigation documents that duplicate the TOC."""
    item_id = str(getattr(item, "id", "") or "").lower()
    item_name = str(item.get_name() or "").replace("\\", "/").lower()
    properties = {str(prop).lower() for prop in getattr(item, "properties", []) or []}
    return item_id == "nav" or item_name.endswith("/nav.xhtml") or item_name == "nav.xhtml" or "nav" in properties


def _replace_markdown_image_alt(link: str, alt: str) -> str:
    """Replace alt text only for standard Markdown image links."""
    if not alt or not link.startswith("![") or link.startswith("![["):
        return link
    end = link.find("]", 2)
    if end == -1:
        return link
    return f"![{alt}]" + link[end + 1 :]
