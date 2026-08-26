"""PDF invoice parsing, scan detection, and page rendering for OCR fallback."""

from __future__ import annotations

from typing import Any, cast

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import _compact_text

MIN_TEXT_LENGTH_FOR_INVOICE = 20
"""Minimum compact text length to consider a PDF page a text-based invoice.

Pages with less than this many non-whitespace / non-zero-width characters
are treated as scanned images and routed to OCR.
"""


def is_scanpage(text: str) -> bool:
    """Return ``True`` when *text* has too few characters to be a text-based invoice.

    Scan-based PDF pages contain no extractable text — PyMuPDF returns
    ``""`` or only noise.  When ``_compact_text(text)`` is shorter than
    ``MIN_TEXT_LENGTH_FOR_INVOICE`` (20 chars), the page should be rendered
    to an image and processed via OCR.
    """
    return len(_compact_text(text)) < MIN_TEXT_LENGTH_FOR_INVOICE


def render_pdf_page_to_png(*, file_path: str, page_index: int, png_path: str) -> str:
    """Render a PDF page to a preprocessed PNG optimised for OCR.

    The pipeline:

    1. Render the page at 300 DPI via PyMuPDF.
    2. Apply EXIF transpose, grayscale, autocontrast.
    3. Sharpen with UnsharpMask (radius=2, percent=180, threshold=3).
    4. Enhance contrast by 1.6x.
    5. Apply adaptive (Otsu) binarisation to produce pure black/white output.

    The PIL processing step is best-effort — if Pillow is unavailable or any
    step raises, the raw rendered PNG is used as-is.

    Args:
        file_path: Path to the PDF file.
        page_index: Zero-based page index.
        png_path: Destination path for the rendered PNG.

    Returns:
        *png_path* (for chaining).
    """
    import fitz

    with fitz.open(file_path, filetype="pdf") as doc:
        page = doc[page_index]
        dpi = 300
        zoom = max(72, int(dpi)) / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        pix.save(png_path)
        try:
            from PIL import Image, ImageEnhance, ImageFilter, ImageOps

            with Image.open(png_path) as src_img:
                img = ImageOps.exif_transpose(src_img)
                img = img.convert("L")
                img = ImageOps.autocontrast(img)
                img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=180, threshold=3))
                img = ImageEnhance.Contrast(img).enhance(1.6)

                # Otsu binarisation — adaptive histogram-based thresholding
                hist = img.histogram()
                total = sum(hist)
                if total > 0:
                    sum_total = 0.0
                    for i, h in enumerate(hist):
                        sum_total += float(i * h)
                    sum_b = 0.0
                    w_b = 0
                    max_var = -1.0
                    threshold = 160
                    for i in range(256):
                        w_b += int(hist[i])
                        if w_b == 0:
                            continue
                        w_f = total - w_b
                        if w_f == 0:
                            break
                        sum_b += float(i * hist[i])
                        m_b = sum_b / float(w_b)
                        m_f = (sum_total - sum_b) / float(w_f)
                        var_between = float(w_b) * float(w_f) * (m_b - m_f) * (m_b - m_f)
                        if var_between > max_var:
                            max_var = var_between
                            threshold = i

                    t = max(90, min(210, int(threshold)))
                    table = [255 if i >= t else 0 for i in range(256)]
                    img = img.point(table, mode="L")
                    img.save(png_path, format="PNG", optimize=True)
        except Exception:
            pass
        return png_path


def read_pdf_text_and_spans(
    file_path: str,
) -> tuple[str, list[tuple[float, float, float, float, str]]]:
    """Extract full text and per-span bounding boxes from all PDF pages.

    Args:
        file_path: Path to the PDF file.

    Returns:
        A tuple of ``(full_text, spans)`` where *spans* is a list of
        ``(x0, y0, x1, y1, text)`` tuples.
    """
    import fitz

    with fitz.open(file_path, filetype="pdf") as doc:
        text_parts: list[str] = []
        spans: list[tuple[float, float, float, float, str]] = []
        for page in doc:
            text_parts.append(str(page.get_text("text")))
            d = cast(dict[str, Any], page.get_text("dict"))
            for block in d.get("blocks") or []:
                if not isinstance(block, dict):
                    continue
                for line in block.get("lines") or []:
                    if not isinstance(line, dict):
                        continue
                    for span in line.get("spans") or []:
                        if not isinstance(span, dict):
                            continue
                        s = (span.get("text") or "").strip()
                        if not s:
                            continue
                        x0, y0, x1, y1 = span.get("bbox") or (0.0, 0.0, 0.0, 0.0)
                        spans.append((float(x0), float(y0), float(x1), float(y1), s))
        return "\n".join(text_parts), spans


def read_pdf_text_and_spans_single_page(
    file_path: str, page_index: int
) -> tuple[str, list[tuple[float, float, float, float, str]]]:
    """Extract text and spans from a single PDF page.

    Args:
        file_path: Path to the PDF file.
        page_index: Zero-based page index.

    Returns:
        A tuple of ``(page_text, spans)``.
    """
    import fitz

    with fitz.open(file_path, filetype="pdf") as doc:
        page = doc[page_index]
        text = str(page.get_text("text"))
        spans: list[tuple[float, float, float, float, str]] = []
        d = cast(dict[str, Any], page.get_text("dict"))
        for block in d.get("blocks") or []:
            if not isinstance(block, dict):
                continue
            for line in block.get("lines") or []:
                if not isinstance(line, dict):
                    continue
                for span in line.get("spans") or []:
                    if not isinstance(span, dict):
                        continue
                    s = (span.get("text") or "").strip()
                    if not s:
                        continue
                    x0, y0, x1, y1 = span.get("bbox") or (0.0, 0.0, 0.0, 0.0)
                    spans.append((float(x0), float(y0), float(x1), float(y1), s))
        return text, spans


def parse_pdf_invoice(
    file_path: str,
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    """Parse a PDF invoice into metadata and detail-line rows.

    This is the main entry point for PDF invoice parsing.
    It reads all pages and passes text+spans to the metadata/rows parser.

    Args:
        file_path: Path to the PDF file.

    Returns:
        A tuple of ``(metadata_dict, rows_list)``.
    """
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.metadata import (
        parse_invoice_metadata_from_text_and_spans,
    )

    text, spans = read_pdf_text_and_spans(file_path)
    return parse_invoice_metadata_from_text_and_spans(text, spans)


def get_pdf_page_count(file_path: str) -> int:
    """Return the number of pages in a PDF file."""
    import fitz

    with fitz.open(file_path, filetype="pdf") as doc:
        return len(doc)
