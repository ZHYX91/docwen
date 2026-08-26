"""Shared image → Markdown link generation.

Works with file paths (not plugin-specific structures) so every plugin can
use it regardless of its internal image representation.
"""

from __future__ import annotations

import base64
import logging
from io import BytesIO
from pathlib import Path

from PIL import Image

from docwen_core.detection import detect_content_format
from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_core.formats import CATEGORY_IMAGE, get_category, get_media_type
from docwen_core.markdown_utils import format_md_file_link
from docwen_core.yaml_tools import generate_basic_yaml_frontmatter

logger = logging.getLogger(__name__)


def build_base64_image_data_uri(
    *,
    image_path: str | Path,
    media_type: str | None = None,
    export_semantics: MarkdownExportSemantics,
) -> str:
    """Return a configured Base64 data URI for an image file.

    Images larger than the configured binary-size threshold are best-effort
    encoded as JPEG.  Compression failures, non-shrinking results, and files
    at or below the threshold preserve the exact source bytes.

    The request-owned *export_semantics* is mandatory: Base64 generation must
    never read process-global configuration or invent a second default source.
    """
    if not isinstance(export_semantics, MarkdownExportSemantics):
        raise TypeError("export_semantics must be a MarkdownExportSemantics instance")
    path = Path(image_path)
    resolved_media_type = _resolve_image_media_type(path, media_type)
    source_bytes = path.read_bytes()
    payload = source_bytes

    compress_enabled = export_semantics.export_base64_compress_enabled
    threshold_kb = export_semantics.export_base64_compress_threshold_kb
    threshold_bytes = threshold_kb * 1024
    if compress_enabled and len(source_bytes) > threshold_bytes:
        try:
            compressed = _compress_image_to_jpeg_bytes(path, threshold_bytes)
        except Exception as exc:
            logger.warning(
                "Base64 image compression failed for %s; using original bytes: %s",
                path.name,
                exc,
            )
        else:
            if len(compressed) >= len(source_bytes):
                logger.warning(
                    "Base64 image compression did not reduce %s; using original bytes (%d >= %d)",
                    path.name,
                    len(compressed),
                    len(source_bytes),
                )
            else:
                payload = compressed
                resolved_media_type = "image/jpeg"
                if len(payload) > threshold_bytes:
                    logger.warning(
                        "Base64 image remains above configured threshold after compression: %s (%d > %d bytes)",
                        path.name,
                        len(payload),
                        threshold_bytes,
                    )

    encoded = base64.b64encode(payload).decode("ascii")
    return f"data:{resolved_media_type};base64,{encoded}"


def _resolve_image_media_type(path: Path, media_type: str | None) -> str:
    detected_format = detect_content_format(str(path)).format
    if get_category(detected_format) != CATEGORY_IMAGE:
        raise ValueError(f"File content is not a supported image: {path}")
    try:
        with Image.open(path) as image:
            image.load()
    except Exception as exc:
        raise ValueError(f"Image content is corrupt or unsupported: {path}") from exc

    resolved = get_media_type(detected_format)
    if media_type:
        declared = media_type.partition(";")[0].strip().lower()
        if declared == "image/jpg":
            declared = "image/jpeg"
        elif declared == "image/tif":
            declared = "image/tiff"
        if declared != resolved:
            logger.warning(
                "Ignoring declared image media type for %s: %s != detected %s",
                path.name,
                declared,
                resolved,
            )
    return resolved


def _compress_image_to_jpeg_bytes(path: Path, target_bytes: int) -> bytes:
    """Encode *path* as the highest-quality JPEG within *target_bytes*.

    When quality 15 is still above the target it is returned so the caller
    can retain the old best-effort behavior and emit an explicit warning.
    """
    working: Image.Image | None = None
    try:
        with Image.open(path) as source:
            if source.mode in {"RGBA", "LA"} or (source.mode == "P" and "transparency" in source.info):
                rgba = source.convert("RGBA")
                try:
                    working = Image.new("RGB", rgba.size, (255, 255, 255))
                    with rgba.getchannel("A") as alpha:
                        working.paste(rgba, mask=alpha)
                finally:
                    rgba.close()
            elif source.mode != "RGB":
                working = source.convert("RGB")
            else:
                working = source.copy()

        quality_min = 15
        quality_max = 95
        best_payload: bytes | None = None
        while quality_min <= quality_max:
            quality = (quality_min + quality_max) // 2
            with BytesIO() as buffer:
                working.save(buffer, format="JPEG", quality=quality, optimize=True)
                candidate = buffer.getvalue()
            if len(candidate) <= target_bytes:
                best_payload = candidate
                quality_min = quality + 1
            else:
                quality_max = quality - 1

        if best_payload is not None:
            return best_payload

        with BytesIO() as buffer:
            working.save(buffer, format="JPEG", quality=15, optimize=True)
            return buffer.getvalue()
    finally:
        if working is not None:
            working.close()


def generate_image_markdown(
    *,
    image_path: str | Path,
    image_mode: str = "file",
    image_link_style: str = "wiki_embed",
    alt_text: str = "",
    export_semantics: MarkdownExportSemantics | None = None,
) -> str:
    """Generate a Markdown image reference from a file path.

    image_mode: ``"file"`` | ``"base64"`` | ``"embed"`` | ``"omit"``.
    image_link_style: ``"wiki_embed"`` | ``"wiki_link"`` |
        ``"markdown_embed"`` | ``"markdown_link"``.
    export_semantics: Request-scoped Base64 compression policy. It is required
        when ``image_mode`` is ``"base64"``.
    """
    if image_mode == "omit":
        return f"<!-- image omitted: {alt_text or Path(image_path).name} -->"
    if image_mode == "base64":
        if export_semantics is None:
            raise TypeError("export_semantics is required for Base64 Markdown images")
        target = build_base64_image_data_uri(
            image_path=image_path,
            export_semantics=export_semantics,
        )
    elif image_mode == "embed":
        # Match the image plugin's local-placeholder semantics: keep a
        # relative reference rather than emitting a data URI.
        target = f"./{Path(image_path).name}"
    else:
        target = str(image_path)
    return _format_link(alt_text, target, image_link_style)


def _format_link(alt: str, target: str, style: str) -> str:
    if style == "wiki_embed":
        return f"![[{target}]]"
    if style == "wiki_link":
        return f"[[{alt or target}|{target}]]"
    if style == "markdown_embed":
        return f"![{alt}]({target})"
    if style == "markdown_link":
        return f"[{alt or target}]({target})"
    return f"![{alt}]({target})"


def build_image_ocr_sidecar(
    *,
    sidecar_stem: str,
    source_format: str,
    image_markdown: str,
    ocr_text: str,
    md_link_style: str,
    ocr_blockquote_title: str | None = None,
    yaml_key_labels: object | None = None,
) -> tuple[str, str]:
    """Build a sidecar .md file for one image with its OCR text.

    This is a **pure function** — it does not touch the filesystem, run OCR,
    register artifacts, or depend on any plugin internals.  Callers are
    responsible for writing the returned sidecar text to disk and for
    plugging the replacement link into the main document.

    Args:
        sidecar_stem: File name stem for the sidecar (no extension).
        source_format: Source format string for YAML front matter.
        image_markdown: The fully-resolved image link as it should appear
            in the sidecar (e.g. ``![[foo.png]]``).
        ocr_text: Raw OCR output text.
        md_link_style: Link style for the main-document replacement link
            (``wiki_embed``, ``wiki_link``, ``markdown_embed``, ``markdown_link``).
        ocr_blockquote_title: Optional bold title line prepended to the
            OCR blockquote (e.g. ``OCR结果`` → ``> **OCR结果**``).

    Returns:
        sidecar_markdown_text: Full sidecar ``.md`` content (front matter +
            image link + OCR blockquote).
        main_doc_replacement_link: The ``.md`` file link that should replace
            the image in the main document (generated via
            :func:`format_md_file_link`).
    """
    sidecar_filename = f"{sidecar_stem}.md"

    # ── YAML front matter ──────────────────────────────────────────
    front_matter = generate_basic_yaml_frontmatter(
        sidecar_stem,
        extra={
            "source_format": source_format,
            "ocr": True,
        },
        yaml_key_labels=yaml_key_labels,
    )

    # ── OCR blockquote ─────────────────────────────────────────────
    ocr_lines: list[str] = []
    if ocr_blockquote_title:
        ocr_lines.append(f"> **{ocr_blockquote_title}**")
        ocr_lines.append(">")
    for line in ocr_text.splitlines():
        stripped = line.strip()
        if not stripped:
            ocr_lines.append(">")
        else:
            ocr_lines.append(f"> {stripped}")
    ocr_block = "\n".join(ocr_lines) + "\n"

    # ── Assemble sidecar content ───────────────────────────────────
    # Structure: front matter → empty line → image link → empty line →
    #            OCR blockquote → trailing empty line
    sidecar_text = front_matter + image_markdown.rstrip("\n") + "\n\n" + ocr_block + "\n"

    # ── Main-document replacement link ─────────────────────────────
    replacement_link = format_md_file_link(sidecar_filename, style=md_link_style)

    return sidecar_text, replacement_link
