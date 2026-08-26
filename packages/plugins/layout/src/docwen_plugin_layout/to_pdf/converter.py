"""Layout/OFD/XPS → PDF converters.

- LayoutToPdfConverter: unified entry for layout→pdf (dispatches by actual format)
- OfdToPdfConverter: direct OFD → PDF (via easyofd)
- XpsToPdfConverter: direct XPS → PDF (via PyMuPDF)
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.paths import input_stem
from docwen_plugin_layout._common import (
    file_size,
    new_artifact_id,
    request_source_format,
)

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext


def _build_pdf_result(
    task_id: str,
    output_path: str,
    stem: str,
    source_format: str,
    context: ConverterContext,
) -> ConversionResult:
    """Build a successful ConversionResult for PDF output."""
    from docwen_core.models.artifact import (
        ARTIFACT_KIND_PRIMARY,
        ArtifactManifest,
    )
    from docwen_core.models.result import (
        ConversionDiagnostic,
        ConversionMetrics,
        ConversionResult,
    )

    out_bytes = file_size(output_path)
    artifact = ArtifactManifest(
        artifact_id=new_artifact_id(),
        kind=ARTIFACT_KIND_PRIMARY,
        staging_path=output_path,
        suggested_name=f"{stem}.pdf",
        media_type="application/pdf",
        metadata={"source_format": source_format},
        is_primary=True,
    )
    context.workspace.add_artifact(artifact)
    context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)

    return ConversionResult(
        task_id=task_id,
        success=True,
        artifacts=[artifact],
        diagnostics=[
            ConversionDiagnostic(
                level="info",
                message=f"Converted {stem} ({source_format}) to PDF ({out_bytes} bytes)",
                code="PDF-CONVERT-OK",
            )
        ],
        metrics=ConversionMetrics(
            input_bytes=file_size(context.workspace.input_path),
            output_bytes=out_bytes,
            extra={"source_format": source_format},
        ),
    )


def _error_result(
    task_id: str,
    message: str,
    code: str,
    exc: Exception | None = None,
    *,
    error_type: str = "conversion_failed",
) -> ConversionResult:
    """Build an error ConversionResult."""
    from docwen_core.models.result import (
        ConversionDiagnostic,
        ConversionErrorInfo,
        ConversionResult,
    )

    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type=error_type,
            message=message,
            diagnostic_code=code,
        ),
        diagnostics=[ConversionDiagnostic(level="error", message=message, code=code)],
    )


# ═══════════════════════════════════════════════════════════════════════
# LayoutToPdfConverter — unified category-level route
# ═══════════════════════════════════════════════════════════════════════


class LayoutToPdfConverter:
    """Handle layout-format PDF output routes.

    Dispatches based on the detected source format:
    - PDF: already a PDF — copies to staging as-is
    - OFD: converts via easyofd
    - XPS: converts via PyMuPDF
    """

    def convert(self, context: ConverterContext) -> ConversionResult:
        task_id = context.request.request_id
        input_path = context.workspace.input_path
        stem = input_stem(input_path)

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting layout → PDF conversion")

        fmt = request_source_format(context)

        try:
            if fmt == "pdf":
                return self._copy_pdf(input_path, task_id, stem, context)
            elif fmt == "ofd":
                return self._ofd_to_pdf(input_path, task_id, stem, context)
            elif fmt == "xps":
                return self._xps_to_pdf(input_path, task_id, stem, context)
            else:
                return self._unsupported_source(task_id, fmt)
        except Exception as exc:
            context.logger.error(f"Layout → PDF failed: {exc}")
            return _error_result(task_id, str(exc), "LAYOUT2PDF-ERROR", exc)

    @staticmethod
    def _copy_pdf(input_path: str, task_id: str, stem: str, context: ConverterContext) -> ConversionResult:
        """PDF is already PDF — copy to staging."""
        output_path = context.workspace.create_artifact_path("primary", ".pdf")
        import shutil

        shutil.copy2(input_path, output_path)
        context.progress.report_progress(100.0, "PDF already in target format")
        return _build_pdf_result(task_id, output_path, stem, "pdf", context)

    @staticmethod
    def _ofd_to_pdf(input_path: str, task_id: str, stem: str, context: ConverterContext) -> ConversionResult:
        """Convert OFD → PDF using easyofd."""
        from docwen_core.ofd import apply_easyofd_patches, easyofd_import_boundary

        try:
            with easyofd_import_boundary():
                from easyofd import OFD  # type: ignore[import-untyped]
        except ImportError:
            msg = "easyofd is not installed. Install it with: pip install easyofd"
            context.logger.error(msg)
            return _error_result(task_id, msg, "OFD2PDF-DEPENDENCY-MISSING")

        apply_easyofd_patches()

        context.cancellation.check()
        context.progress.report_progress(30.0, "Converting OFD to PDF")

        ofd = OFD()
        ofd.read(input_path, fmt="path")

        context.cancellation.check()
        pdf_bytes = ofd.to_pdf()

        output_path = context.workspace.create_artifact_path("primary", ".pdf")
        Path(output_path).write_bytes(pdf_bytes)

        context.progress.report_progress(100.0, "OFD → PDF complete")
        return _build_pdf_result(task_id, output_path, stem, "ofd", context)

    @staticmethod
    def _xps_to_pdf(input_path: str, task_id: str, stem: str, context: ConverterContext) -> ConversionResult:
        """Convert XPS → PDF using PyMuPDF."""
        try:
            import fitz  # PyMuPDF
        except ImportError:
            msg = "PyMuPDF is not installed. Install it with: pip install PyMuPDF"
            context.logger.error(msg)
            return _error_result(task_id, msg, "XPS2PDF-DEPENDENCY-MISSING")

        context.cancellation.check()
        context.progress.report_progress(30.0, "Converting XPS to PDF")

        with fitz.open(input_path, filetype="xps") as doc:
            pdf_bytes = doc.convert_to_pdf()

        context.cancellation.check()

        with fitz.open("pdf", pdf_bytes) as pdf_doc:
            output_path = context.workspace.create_artifact_path("primary", ".pdf")
            pdf_doc.save(output_path)

        context.progress.report_progress(100.0, "XPS → PDF complete")
        return _build_pdf_result(task_id, output_path, stem, "xps", context)

    @staticmethod
    def _unsupported_source(task_id: str, source_format: str) -> ConversionResult:
        msg = f"Unsupported admitted layout source format: {source_format or 'unknown'}."
        return _error_result(task_id, msg, "LAYOUT2PDF-UNSUPPORTED-SOURCE", error_type="invalid_input")


# ═══════════════════════════════════════════════════════════════════════
# OfdToPdfConverter — direct OFD→PDF route
# ═══════════════════════════════════════════════════════════════════════


class OfdToPdfConverter:
    """Direct OFD → PDF converter (mirrors the category-level dispatcher)."""

    def convert(self, context: ConverterContext) -> ConversionResult:
        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting OFD → PDF conversion")
        return LayoutToPdfConverter()._ofd_to_pdf(
            context.workspace.input_path,
            context.request.request_id,
            input_stem(context.workspace.input_path),
            context,
        )


# ═══════════════════════════════════════════════════════════════════════
# XpsToPdfConverter — direct XPS→PDF route
# ═══════════════════════════════════════════════════════════════════════


class XpsToPdfConverter:
    """Direct XPS → PDF converter (mirrors the category-level dispatcher)."""

    def convert(self, context: ConverterContext) -> ConversionResult:
        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting XPS → PDF conversion")
        return LayoutToPdfConverter()._xps_to_pdf(
            context.workspace.input_path,
            context.request.request_id,
            input_stem(context.workspace.input_path),
            context,
        )
