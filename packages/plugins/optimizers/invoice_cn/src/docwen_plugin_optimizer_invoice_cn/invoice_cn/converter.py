"""InvoiceCnConverter — converts PDF/OFD/image invoices to structured Markdown.

Produces YAML frontmatter (20 metadata fields) + a Markdown detail-line table.
For image-based invoices, OCR is applied first to extract text.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.text.ocr import OcrOutcome, OcrStatus, format_ocr_best_effort_warning

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext


def _ocr_and_parse_image_invoice(
    input_path: str,
    *,
    source_format: str,
    ocr_language: str = "auto",
    current_locale: str = "zh_CN",
) -> tuple[dict[str, str | None], list[dict[str, str]], OcrOutcome]:
    """OCR an image invoice and parse the recognised text into metadata + rows.

    Uses the same compact-text + row parsing pipeline as PDF invoices.
    """
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.image_parser import (
        parse_image_invoice_outcome,
    )

    return parse_image_invoice_outcome(
        input_path,
        source_format=source_format,
        ocr_language=ocr_language,
        current_locale=current_locale,
    )


def _report_ocr_best_effort(
    context: ConverterContext,
    outcome: OcrOutcome,
    *,
    location: str,
) -> None:
    message = format_ocr_best_effort_warning(outcome.status)
    if message is None:
        return
    context.logger.warning(f"{message} {location}: {outcome.message}".rstrip())
    context.progress.report_diagnostic(
        "warning",
        message,
        code="OCR-BEST-EFFORT",
        location=location,
    )


class InvoiceCnConverter:
    """Convert a Chinese invoice (PDF/OFD/image) to structured Markdown.

    Supports:
    - PDF invoices via PyMuPDF text+spans extraction + regex parsing
    - OFD invoices via InvoiceData.xml (or fallback content.xml extraction)
    - Image invoices via OCR + regex parsing
    """

    def convert(self, context: ConverterContext) -> ConversionResult:
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
        from docwen_plugin_optimizer_invoice_cn._common import (
            file_size,
            new_artifact_id,
            request_source_format,
        )
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.markdown_writer import (
            _build_yaml_frontmatter,
            _render_markdown_table,
        )
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import (
            INVOICE_CN_YAML_SCHEMA,
            TABLE_HEADERS,
        )

        task_id = context.request.request_id
        input_path = context.workspace.input_path
        options = context.request.options
        input_stem_val = input_stem(input_path)
        source_format = request_source_format(context)

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting invoice to Markdown conversion")

        enable_ocr = bool(options.get("to_md_enable_ocr", False))
        ocr_language = str(options.get("ocr_language") or "auto")
        current_locale = str(options.get("locale") or "zh_CN")
        ocr_succeeded = False
        scan_pages: list[int] = []  # Track scan-detected pages for diagnostics

        # ── Image source — requires OCR ───────────────────────────────
        if source_format in ("jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp", "heic", "heif"):
            context.progress.report_progress(10.0, "Running OCR on image invoice")

            # OCR the image and parse
            try:
                metadata, rows, outcome = _ocr_and_parse_image_invoice(
                    input_path,
                    source_format=source_format,
                    ocr_language=ocr_language,
                    current_locale=current_locale,
                )
            except Exception as exc:
                metadata, rows = {}, []
                outcome = OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))
            _report_ocr_best_effort(
                context,
                outcome,
                location=Path(input_path).name,
            )
            ocr_succeeded = outcome.status is OcrStatus.SUCCESS

        else:
            # ── PDF / OFD source ──────────────────────────────────────
            context.cancellation.check()
            context.progress.report_progress(10.0, "Parsing invoice file")

            metadata: dict[str, str | None] = {}
            rows: list[dict[str, str]] = []

            if source_format == "pdf":
                context.progress.report_progress(20.0, "Extracting PDF invoice text and spans")

                from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
                    get_pdf_page_count,
                    is_scanpage,
                    read_pdf_text_and_spans_single_page,
                    render_pdf_page_to_png,
                )

                # ── Scan-page detection (outside OCR gate) ───────────────
                page_count = get_pdf_page_count(input_path)
                scan_pages.clear()
                for page_idx in range(page_count):
                    try:
                        page_text, _ = read_pdf_text_and_spans_single_page(input_path, page_idx)
                    except Exception:
                        page_text = ""
                    if is_scanpage(page_text):
                        scan_pages.append(page_idx)

                # ── OCR fallback for scan pages ──────────────────────────
                if scan_pages and enable_ocr:
                    context.progress.report_progress(
                        25.0,
                        "Scan-based PDF page detected — rendering to PNG for OCR",
                    )
                    import tempfile
                    from pathlib import Path as _Path

                    try:
                        with tempfile.TemporaryDirectory() as tmpdir:
                            png_path = str(_Path(tmpdir) / "__page_0.png")
                            render_pdf_page_to_png(
                                file_path=input_path,
                                page_index=scan_pages[0],
                                png_path=png_path,
                            )
                            metadata, rows, outcome = _ocr_and_parse_image_invoice(
                                png_path,
                                source_format="png",
                                ocr_language=ocr_language,
                                current_locale=current_locale,
                            )
                    except Exception as exc:
                        outcome = OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))

                    if outcome.status is not OcrStatus.SUCCESS:
                        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
                            parse_pdf_invoice,
                        )

                        metadata, rows = parse_pdf_invoice(input_path)
                    _report_ocr_best_effort(
                        context,
                        outcome,
                        location=f"{Path(input_path).name}:page-{scan_pages[0] + 1}",
                    )
                    ocr_succeeded = outcome.status is OcrStatus.SUCCESS

                elif scan_pages and not enable_ocr:
                    # Warn about scan pages but still parse as-is
                    context.progress.report_progress(
                        25.0,
                        "Scan page(s) detected but OCR is disabled — parsing available text",
                    )

                    from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
                        parse_pdf_invoice,
                    )

                    metadata, rows = parse_pdf_invoice(input_path)
                else:
                    from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
                        parse_pdf_invoice,
                    )

                    metadata, rows = parse_pdf_invoice(input_path)

            elif source_format == "ofd":
                context.progress.report_progress(20.0, "Extracting OFD invoice data")
                from docwen_plugin_optimizer_invoice_cn.invoice_cn.ofd_parser import (
                    parse_ofd_invoice,
                )

                metadata, rows = parse_ofd_invoice(input_path)

            else:
                msg = f"Unsupported invoice source format: {source_format}"
                context.logger.error(msg)
                return ConversionResult(
                    task_id=task_id,
                    success=False,
                    error=ConversionErrorInfo(
                        error_type="invalid_input",
                        message=msg,
                        diagnostic_code="INVOICE-INVALID-FORMAT",
                    ),
                    diagnostics=[
                        ConversionDiagnostic(
                            level="error",
                            message=msg,
                            code="INVOICE-INVALID-FORMAT",
                        )
                    ],
                )

        input_bytes = file_size(input_path)

        try:
            context.cancellation.check()

            # ── Build YAML frontmatter ───────────────────────────────
            context.progress.report_progress(60.0, "Building output Markdown")
            metadata_yaml: dict[str, str | None] = {}
            for k in INVOICE_CN_YAML_SCHEMA:
                v = metadata.get(k)
                metadata_yaml[k] = str(v).strip() if v is not None else ""

            yaml_frontmatter = _build_yaml_frontmatter(
                file_stem=input_stem_val,
                metadata=metadata_yaml,
                include_empty=True,
                yaml_key_labels=options.get("yaml_key_labels"),
            )
            table_md = _render_markdown_table(headers=TABLE_HEADERS, rows=rows)
            md_text = yaml_frontmatter + "## 商品明细\n\n" + table_md

            # ── Write Markdown artifact to staging ───────────────────
            md_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".md")
            Path(md_path).write_text(md_text, encoding="utf-8")

            artifact = ArtifactManifest(
                artifact_id=new_artifact_id(),
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=md_path,
                suggested_name=f"{input_stem_val}.md",
                media_type="text/markdown",
                metadata={
                    "source_format": source_format,
                    "row_count": len(rows),
                    "yaml_fields": len(INVOICE_CN_YAML_SCHEMA),
                },
                is_primary=True,
            )
            context.workspace.add_artifact(artifact)
            context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)

            output_bytes = file_size(md_path)
            context.progress.report_progress(100.0, "Invoice to Markdown complete")

            diagnostics = [
                ConversionDiagnostic(
                    level="info",
                    message=(
                        f"Converted {input_stem_val} to invoice Markdown "
                        f"({len(rows)} detail lines, {output_bytes} bytes)"
                    ),
                    code="INVOICE-OK",
                )
            ]
            if ocr_succeeded:
                diagnostics.append(
                    ConversionDiagnostic(
                        level="info",
                        message="OCR-based invoice conversion",
                        code="INVOICE-OCR-OK",
                    )
                )

            if scan_pages:
                diagnostics.append(
                    ConversionDiagnostic(
                        level="warning",
                        message=(
                            f"Page(s) {scan_pages} appear to be scanned images. "
                            f"OCR is {'enabled' if enable_ocr else 'disabled'}."
                        ),
                        code="INVOICE-SCAN-DETECTED",
                    )
                )

            return ConversionResult(
                task_id=task_id,
                success=True,
                artifacts=[artifact],
                diagnostics=diagnostics,
                metrics=ConversionMetrics(
                    input_bytes=input_bytes,
                    output_bytes=output_bytes,
                    extra={
                        "source_format": source_format,
                        "row_count": len(rows),
                    },
                ),
            )

        except Exception as exc:
            context.logger.error(f"Invoice conversion failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="INVOICE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Invoice conversion failed: {exc}",
                        code="INVOICE-ERROR",
                    )
                ],
            )
