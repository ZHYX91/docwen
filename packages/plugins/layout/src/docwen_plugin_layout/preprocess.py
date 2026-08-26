"""Preprocessing layer for fixed-layout inputs.

Handles OFD/XPS -> PDF pre-conversion so that downstream converters
(to_markdown, to_image, to_document) always receive PDF input.
PDF inputs pass through unchanged.

This centralises the old pre-processing chain that used to live in
``services/strategies/layout/`` so that the Layout plugin itself is
responsible for making non-PDF fixed-layout sources reachable.

Also provides HTML image preprocessing (build_image_markdown,
materialize_image_target) for resolving, copying, and formatting
``<img>`` elements into Markdown image links with OCR support,
base-href resolution, and configurable link styling.
"""

from __future__ import annotations

import logging
import shutil
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse
from urllib.request import url2pathname

logger = logging.getLogger(__name__)


@dataclass
class PreprocessResult:
    """Result of preprocessing a fixed-layout input.

    On success the caller should switch to *effective_input_path* and
    *effective_source_format* before handing off to a converter.

    On failure *error_type* / *error_message* / *diagnostic_code* are set
    and the caller should return the appropriate error result.
    """

    effective_input_path: str
    """Path to the file that downstream converters should read."""

    original_source_format: str
    """The source format the caller originally asked for (``"ofd"``, …)."""

    effective_source_format: str
    """The format the effective input actually is (``"pdf"`` after pre-processing)."""

    intermediate_artifacts: list[str] = field(default_factory=list)
    """Paths to intermediate files created during preprocessing.

    These are NOT registered as workspace artifacts — they are internal
    staging files that may be cleaned up later.
    """

    # -- error fields (set when preprocessing fails) -----------------------

    error_type: str | None = None
    """One of ``"invalid_input"``, ``"dependency_missing"``, ``"conversion_failed"``."""

    error_message: str | None = None
    """Human-readable error description."""

    diagnostic_code: str | None = None
    """Machine-readable diagnostic code."""


def preprocess_layout_input(
    input_path: str,
    staging_dir: str,
    source_format: str,
) -> PreprocessResult:
    """Preprocess a fixed-layout input, converting OFD/XPS to PDF if needed.

    Args:
        input_path: Absolute path to the source file.
        staging_dir: Writable directory for intermediate files.
        source_format: The admitted source format (``"pdf"``, ``"ofd"``, or ``"xps"``).

    Returns:
        A ``PreprocessResult`` — check ``error_type`` before using the
        effective path / format.
    """
    fmt = source_format.lower().strip()

    # ``FileRef.format`` must be concrete after admission. Refuse a generic
    # category instead of guessing again from a possibly misleading suffix.
    if fmt == "layout":
        return PreprocessResult(
            effective_input_path=input_path,
            original_source_format=source_format,
            effective_source_format=source_format,
            error_type="invalid_input",
            error_message="Layout preprocessing requires a concrete admitted source format.",
            diagnostic_code="PREPROCESS-SOURCE-FORMAT-NOT-CONCRETE",
        )

    # ── PDF: pass-through ────────────────────────────────────────────────
    if fmt == "pdf":
        return PreprocessResult(
            effective_input_path=input_path,
            original_source_format=source_format,
            effective_source_format="pdf",
        )

    # ── OFD → PDF ────────────────────────────────────────────────────────
    if fmt == "ofd":
        return _ofd_to_pdf(input_path, staging_dir)

    # ── XPS → PDF ────────────────────────────────────────────────────────
    if fmt == "xps":
        return _xps_to_pdf(input_path, staging_dir)

    # ── Unknown: fail closed; parser selection is an admission decision ───
    return PreprocessResult(
        effective_input_path=input_path,
        original_source_format=source_format,
        effective_source_format=source_format,
        error_type="invalid_input",
        error_message=f"Unsupported admitted layout source format: {source_format or 'unknown'}.",
        diagnostic_code="PREPROCESS-SOURCE-FORMAT-UNSUPPORTED",
    )


# ═══════════════════════════════════════════════════════════════════════════
# Internal helpers
# ═══════════════════════════════════════════════════════════════════════════


def _ofd_to_pdf(input_path: str, staging_dir: str) -> PreprocessResult:
    """Convert OFD → PDF using easyofd.

    Returns a *PreprocessResult* whose *effective_source_format* is
    ``"pdf"`` on success.
    """
    from docwen_core.ofd import apply_easyofd_patches, easyofd_import_boundary

    try:
        with easyofd_import_boundary():
            from easyofd import OFD  # type: ignore[import-untyped]
    except ImportError:
        return PreprocessResult(
            effective_input_path=input_path,
            original_source_format="ofd",
            effective_source_format="ofd",
            error_type="dependency_missing",
            error_message=(
                "easyofd is not installed — required for OFD→PDF preprocessing. Install with: pip install easyofd"
            ),
            diagnostic_code="OFD2PDF-DEPENDENCY-MISSING",
        )

    # Apply runtime monkey-patches before using easyofd (F-I2a-003).
    # - Redirects FileRead scratch files from CWD to tempdir.
    # - Wraps draw_annotation with per-item error isolation.
    apply_easyofd_patches()

    try:
        ofd = OFD()
        ofd.read(input_path, fmt="path")
        pdf_bytes = ofd.to_pdf()
    except Exception as exc:
        return PreprocessResult(
            effective_input_path=input_path,
            original_source_format="ofd",
            effective_source_format="ofd",
            error_type="conversion_failed",
            error_message=f"OFD → PDF preprocessing failed: {exc}",
            diagnostic_code="OFD2PDF-ERROR",
        )

    output_path = str(Path(staging_dir) / f"_preprocess_ofd_{uuid.uuid4().hex[:8]}.pdf")
    Path(output_path).write_bytes(pdf_bytes)

    return PreprocessResult(
        effective_input_path=output_path,
        original_source_format="ofd",
        effective_source_format="pdf",
        intermediate_artifacts=[output_path],
    )


def _xps_to_pdf(input_path: str, staging_dir: str) -> PreprocessResult:
    """Convert XPS → PDF using PyMuPDF.

    Returns a *PreprocessResult* whose *effective_source_format* is
    ``"pdf"`` on success.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return PreprocessResult(
            effective_input_path=input_path,
            original_source_format="xps",
            effective_source_format="xps",
            error_type="dependency_missing",
            error_message=(
                "PyMuPDF (fitz) is not installed — required for XPS→PDF preprocessing. "
                "Install with: pip install PyMuPDF"
            ),
            diagnostic_code="XPS2PDF-DEPENDENCY-MISSING",
        )

    try:
        with fitz.open(input_path, filetype="xps") as doc:  # type: ignore[possibly-undefined]
            pdf_bytes = doc.convert_to_pdf()
    except Exception as exc:
        return PreprocessResult(
            effective_input_path=input_path,
            original_source_format="xps",
            effective_source_format="xps",
            error_type="conversion_failed",
            error_message=f"XPS → PDF preprocessing failed: {exc}",
            diagnostic_code="XPS2PDF-ERROR",
        )

    output_path = str(Path(staging_dir) / f"_preprocess_xps_{uuid.uuid4().hex[:8]}.pdf")
    with fitz.open("pdf", pdf_bytes) as pdf_doc:  # type: ignore[possibly-undefined]
        pdf_doc.save(output_path)

    return PreprocessResult(
        effective_input_path=output_path,
        original_source_format="xps",
        effective_source_format="pdf",
        intermediate_artifacts=[output_path],
    )


# ═══════════════════════════════════════════════════════════════════════════════
# HTML image Markdown preprocessing  (F-G3-004, F-G3-005, F-G3-007)
# ═══════════════════════════════════════════════════════════════════════════════


def _is_remote_url(value: str) -> bool:
    """Return ``True`` when *value* is an ``http`` or ``https`` URL."""
    p = urlparse(value)
    return p.scheme in {"http", "https"}


def _extract_base_href(html_text: str) -> str | None:
    """Extract the ``<base href="...">`` value from raw HTML text.

    Uses a regex so the caller does not need lxml just for base-href
    extraction.  Returns ``None`` when no ``<base>`` tag is found.
    """
    import re

    m = re.search(
        r'<base\s[^>]*href\s*=\s*["\']([^"\']*)["\'][^>]*/?\s*>',
        html_text,
        re.IGNORECASE,
    )
    if m is None:
        return None
    href = m.group(1).strip()
    return href or None


def _resolve_local_path(
    *,
    src: str,
    html_path: str,
    base_href: str | None,
    resource_dir: str | None,
) -> Path | None:
    """Resolve an image ``src`` attribute to a local file path.

    Handles ``file://`` URIs, absolute paths, and relative paths resolved
    against the HTML file's directory or an optional *resource_dir*.

    Returns the absolute ``Path`` on success, or ``None`` when the source
    is a non-``file`` scheme or cannot be resolved.
    """
    parsed = urlparse(src)
    if parsed.scheme == "file":
        try:
            uri_path = parsed.path
            if parsed.netloc and parsed.netloc != "localhost":
                uri_path = f"//{parsed.netloc}{uri_path}"
            return Path(url2pathname(uri_path))
        except Exception:
            return None

    if parsed.scheme:
        return None

    raw_path = parsed.path or src
    raw_path = unquote(raw_path)
    raw_path = raw_path.replace("\\", "/")

    base_dir = Path(resource_dir) if resource_dir else Path(html_path).parent

    if raw_path.startswith("/") and base_href and base_href.startswith("file:"):
        try:
            base_parsed = urlparse(base_href)
            base_uri_path = base_parsed.path
            if base_parsed.netloc and base_parsed.netloc != "localhost":
                base_uri_path = f"//{base_parsed.netloc}{base_uri_path}"
            root = Path(url2pathname(base_uri_path)).parent
            candidate = root / raw_path.lstrip("/")
            return candidate
        except Exception:
            return None

    candidate = (base_dir / raw_path).resolve()
    return candidate


def _make_image_filename(
    src_path: Path,
    original_basename: str,
    image_index: int,
    unified_timestamp_desc: str,
) -> str:
    """Build a sanitised output filename for a materialised image.

    Uses ``docwen_core.markdown_utils.sanitize_filename`` so the result
    is safe for every target filesystem.
    """
    from docwen_core.detection import detect_content_format
    from docwen_core.formats import get_category
    from docwen_core.markdown_utils import sanitize_filename

    detected_format = detect_content_format(str(src_path)).format
    if get_category(detected_format) != "image":
        raise ValueError(f"HTML image resource is not a supported image: {src_path}")
    ext = {"jpeg": "jpg", "tiff": "tif"}.get(detected_format, detected_format)
    raw = f"{original_basename}_image{image_index}_{unified_timestamp_desc}.{ext}"
    return sanitize_filename(raw)


def _copy_local_image(
    *,
    src_path: Path,
    output_folder: str,
    original_basename: str,
    image_index: int,
    unified_timestamp_desc: str,
) -> str | None:
    """Copy a local image file to *output_folder* with a structured filename.

    Returns the filename (not full path) of the copied image, or ``None``
    when the resource is not actually an image. Skips the copy when the
    target already exists.
    """
    try:
        filename = _make_image_filename(src_path, original_basename, image_index, unified_timestamp_desc)
    except (OSError, ValueError):
        return None
    target = Path(output_folder) / filename
    if not target.exists():
        shutil.copy2(src_path, target)
    return filename


def _save_data_uri_image(
    *,
    src: str,
    output_folder: str,
    original_basename: str,
    image_index: int,
    unified_timestamp_desc: str,
) -> str:
    """Decode a ``data:`` URI image and save it to *output_folder*.

    Uses the shared ``docwen_core.links`` data-URI utilities for detection
    and temp-file materialisation.

    Returns the filename (not full path) on success, or the original *src*
    when the data URI is unrecognised or decoding fails.
    """
    from docwen_core.links import is_data_uri_image, resolve_data_uri_image_to_temp_file

    if not is_data_uri_image(src):
        return src

    temp_file = resolve_data_uri_image_to_temp_file(src, temp_dir=None)
    if not temp_file:
        return src

    src_path = Path(temp_file)
    try:
        filename = _make_image_filename(src_path, original_basename, image_index, unified_timestamp_desc)
    except (OSError, ValueError):
        return src
    target = Path(output_folder) / filename
    if not target.exists():
        shutil.copy2(src_path, target)
    return filename


def materialize_image_target(
    *,
    src: str,
    html_path: str,
    base_href: str | None,
    resource_dir: str | None,
    output_folder: str,
    original_basename: str,
    image_index: int,
    unified_timestamp_desc: str,
) -> dict[str, str | None]:
    """Materialise an image target to the output folder **without** OCR.

    This is the simpler "just save the image" path.  Returns a dict
    with ``filename`` and ``image_path`` keys describing the materialised
    image, or ``image_path=None`` when the source is a remote URL.

    F-G3-005
    """
    # ── data: URI ──────────────────────────────────────────────────────
    if src.startswith("data:"):
        filename = _save_data_uri_image(
            src=src,
            output_folder=output_folder,
            original_basename=original_basename,
            image_index=image_index,
            unified_timestamp_desc=unified_timestamp_desc,
        )
        image_path = str(Path(output_folder) / filename) if filename != src else None
        return {"filename": filename, "image_path": image_path}

    # ── Remote URL ─────────────────────────────────────────────────────
    if _is_remote_url(src):
        return {"filename": src, "image_path": None}

    # ── base_href urljoin for relative→absolute ────────────────────────
    if base_href and _is_remote_url(base_href):
        try:
            joined = urljoin(base_href, src)
            if _is_remote_url(joined):
                return {"filename": joined, "image_path": None}
        except Exception as exc:
            logger.warning("HTML image target urljoin failed, falling back: %s", exc)

    # ── Local path ─────────────────────────────────────────────────────
    local_path = _resolve_local_path(src=src, html_path=html_path, base_href=base_href, resource_dir=resource_dir)
    if local_path is None or not local_path.exists():
        return {"filename": src, "image_path": None}

    filename = _copy_local_image(
        src_path=local_path,
        output_folder=output_folder,
        original_basename=original_basename,
        image_index=image_index,
        unified_timestamp_desc=unified_timestamp_desc,
    )
    if filename is None:
        return {"filename": src, "image_path": None}
    return {"filename": filename, "image_path": str(Path(output_folder) / filename)}


def build_image_markdown(
    *,
    src: str,
    html_path: str,
    base_href: str | None,
    resource_dir: str | None,
    output_folder: str,
    original_basename: str,
    image_index: int,
    unified_timestamp_desc: str,
    keep_images: bool = True,
    enable_ocr: bool = False,
    image_link_style: str = "wiki_embed",
    md_file_link_style: str = "wiki_embed",
    ocr_blockquote_title: str = "",
    ocr_language: str = "auto",
    current_locale: str = "zh_CN",
) -> str:
    """Convert an HTML ``<img src>`` into a Markdown image link.

    Handles data URIs, remote URLs, and local paths — with optional OCR
    text extraction and configurable link styling.

    Returns a Markdown image link string (e.g. ``![alt](path)`` or
    ``![[path]]``), or an empty string when the source cannot be resolved.

    F-G3-004
    """
    from docwen_core.export_semantics import format_image_link

    # ── data: URI ──────────────────────────────────────────────────────
    if src.startswith("data:"):
        from docwen_core.links import is_data_uri_image, resolve_data_uri_image_to_temp_file

        if not is_data_uri_image(src):
            return ""
        temp_file = resolve_data_uri_image_to_temp_file(src, temp_dir=None)
        if not temp_file:
            return ""
        src_path = Path(temp_file)
        try:
            filename = _make_image_filename(src_path, original_basename, image_index, unified_timestamp_desc)
        except (OSError, ValueError):
            return ""
        if keep_images:
            target = Path(output_folder) / filename
            if not target.exists():
                shutil.copy2(src_path, target)
            target_path = str(target)
        else:
            target_path = str(src_path)

        if enable_ocr:
            return _ocr_image_link(
                image_path=target_path,
                filename=filename,
                keep_images=keep_images,
                output_folder=output_folder,
                image_link_style=image_link_style,
                md_file_link_style=md_file_link_style,
                ocr_blockquote_title=ocr_blockquote_title,
                ocr_language=ocr_language,
                current_locale=current_locale,
            )
        return format_image_link(filename, filename, style=image_link_style)

    # ── Remote URL ─────────────────────────────────────────────────────
    if _is_remote_url(src):
        return format_image_link(src, src, style=image_link_style) if keep_images else ""

    # ── base_href urljoin for relative→absolute ────────────────────────
    if base_href and _is_remote_url(base_href):
        try:
            joined = urljoin(base_href, src)
            if _is_remote_url(joined):
                return format_image_link(joined, joined, style=image_link_style) if keep_images else ""
        except Exception as exc:
            logger.warning("HTML image urljoin failed, falling back: %s", exc)

    # ── Local path ─────────────────────────────────────────────────────
    local_path = _resolve_local_path(src=src, html_path=html_path, base_href=base_href, resource_dir=resource_dir)
    if local_path is None or not local_path.exists():
        return format_image_link(src, src, style=image_link_style) if keep_images else ""

    try:
        filename = _make_image_filename(local_path, original_basename, image_index, unified_timestamp_desc)
    except (OSError, ValueError):
        return format_image_link(src, src, style=image_link_style) if keep_images else ""
    if keep_images:
        target = Path(output_folder) / filename
        if not target.exists():
            shutil.copy2(local_path, target)
        target_path = str(target)
    else:
        target_path = str(local_path)

    if enable_ocr:
        return _ocr_image_link(
            image_path=target_path,
            filename=filename,
            keep_images=keep_images,
            output_folder=output_folder,
            image_link_style=image_link_style,
            md_file_link_style=md_file_link_style,
            ocr_blockquote_title=ocr_blockquote_title,
            ocr_language=ocr_language,
            current_locale=current_locale,
        )
    return format_image_link(filename, filename, style=image_link_style)


def _ocr_image_link(
    *,
    image_path: str,
    filename: str,
    keep_images: bool,
    output_folder: str,
    image_link_style: str,
    md_file_link_style: str,
    ocr_blockquote_title: str,
    ocr_language: str,
    current_locale: str,
) -> str:
    """Build a Markdown link for an image, optionally including OCR text.

    When OCR is enabled and text is recognised, the output includes both
    the image link and the recognised text (as a blockquote when a title
    is configured, or plain text otherwise).

    Uses the shared typed ``docwen_core.text.ocr.run_ocr_outcome`` entry so
    empty recognition and operational failure remain distinguishable.
    """
    from docwen_core.detection import detect_content_format
    from docwen_core.export_semantics import format_image_link
    from docwen_core.text.ocr import OcrOutcome, OcrStatus, format_ocr_best_effort_warning, run_ocr_outcome

    image_link = format_image_link(filename, filename, style=image_link_style)

    if not keep_images:
        return image_link

    try:
        outcome = run_ocr_outcome(
            image_path,
            source_format=detect_content_format(image_path).format,
            ocr_language=ocr_language,
            current_locale=current_locale,
        )
    except Exception as exc:
        outcome = OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))

    if outcome.status is not OcrStatus.SUCCESS:
        if warning := format_ocr_best_effort_warning(outcome.status, context=f"HTML image {filename}"):
            logger.warning("%s %s", warning, outcome.message)
        return image_link
    ocr_text = outcome.recognized_text

    if ocr_blockquote_title:
        return f"{image_link}\n\n> **{ocr_blockquote_title}**\n>\n> " + ocr_text.replace("\n", "\n> ") + "\n"

    return f"{image_link}\n\n{ocr_text}\n"


def preprocess_html_images(
    *,
    html_text: str,
    html_path: str,
    output_folder: str,
    resource_dir: str | None = None,
    keep_images: bool = True,
    enable_ocr: bool = False,
    image_link_style: str = "wiki_embed",
    md_file_link_style: str = "wiki_embed",
    ocr_blockquote_title: str = "",
    unified_timestamp_desc: str = "export",
    ocr_language: str = "auto",
    current_locale: str = "zh_CN",
) -> dict[str, str]:
    """Preprocess ``<img>`` elements in *html_text* for Markdown conversion.

    Iterates every ``<img>`` tag, builds a Markdown image link via
    :func:`build_image_markdown`, and replaces the tag with a token.
    Returns a dict with:

    * ``"html"`` — the modified HTML (``<img>`` tags replaced by tokens).
    * ``"token_map"`` — a JSON-serialisable ``{token: markdown_link}``
      mapping for post-conversion replacement.

    The caller should convert the returned HTML to Markdown (e.g. via
    ``markdownify``) and then replace each token with its value.
    """
    import re

    token_map: dict[str, str] = {}
    original_basename = Path(html_path).stem
    base_href = _extract_base_href(html_text)
    image_index = 0

    def _replace_img(m: re.Match[str]) -> str:
        nonlocal image_index
        img_tag = m.group(0)
        # Extract src attribute
        src_m = re.search(r'src\s*=\s*["\']([^"\']*)["\']', img_tag, re.IGNORECASE)
        if src_m is None:
            return img_tag
        src = src_m.group(1).strip()
        if not src:
            return img_tag

        image_index += 1
        token = f"DOCWENIMG{image_index:04d}"
        md_link = build_image_markdown(
            src=src,
            html_path=html_path,
            base_href=base_href,
            resource_dir=resource_dir,
            output_folder=output_folder,
            original_basename=original_basename,
            image_index=image_index,
            unified_timestamp_desc=unified_timestamp_desc,
            keep_images=keep_images,
            enable_ocr=enable_ocr,
            image_link_style=image_link_style,
            md_file_link_style=md_file_link_style,
            ocr_blockquote_title=ocr_blockquote_title,
            ocr_language=ocr_language,
            current_locale=current_locale,
        )
        token_map[token] = md_link
        return token

    processed = re.sub(r"<img\s[^>]*/?\s*>", _replace_img, html_text, flags=re.IGNORECASE)
    return {"html": processed, "token_map": token_map}  # pyright: ignore[reportReturnType]
