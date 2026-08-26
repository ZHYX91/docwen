"""Image invoice parsing via the shared typed OCR API."""

from __future__ import annotations

from docwen_core.text.ocr import (
    OcrOutcome,
    run_ocr_outcome,
)


def parse_image_invoice_outcome(
    input_path: str,
    *,
    source_format: str,
    ocr_language: str = "auto",
    current_locale: str = "zh_CN",
) -> tuple[dict[str, str | None], list[dict[str, str]], OcrOutcome]:
    """OCR and parse an image invoice while retaining the OCR status."""
    outcome = run_ocr_outcome(
        input_path,
        source_format=source_format,
        ocr_language=ocr_language,
        current_locale=current_locale,
    )
    ocr_text = outcome.recognized_text
    if not ocr_text.strip():
        return {}, [], outcome

    from docwen_plugin_optimizer_invoice_cn.invoice_cn.metadata import (
        parse_invoice_metadata_from_compact_text,
    )
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import _compact_text
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import (
        parse_invoice_rows_from_pdf_text,
    )

    compact = _compact_text(ocr_text)
    metadata = parse_invoice_metadata_from_compact_text(compact)
    rows = parse_invoice_rows_from_pdf_text(ocr_text, prefer_marked=False)

    return metadata, rows, outcome
