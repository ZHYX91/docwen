"""Office bridge-backed document/spreadsheet → PDF converter."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.result import ConversionDiagnostic
from docwen_core.office_bridge import BridgeCandidate, BridgeResult, convert_with_backend_priority
from docwen_plugin_print.to_pdf.converter import _build_pdf_result, _error_result

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext


_DEFAULT_PRIORITIES = {
    "spreadsheet": ("wps_spreadsheets", "msoffice_excel", "libreoffice"),
    "document": ("wps_writer", "msoffice_word", "libreoffice"),
}

_COM_CANDIDATES = {
    "spreadsheet": {
        "wps_spreadsheets": BridgeCandidate("WPS Spreadsheets", "KET.Application", 57, "excel"),
        "msoffice_excel": BridgeCandidate("Microsoft Excel", "Excel.Application", 57, "excel"),
    },
    "document": {
        "wps_writer": BridgeCandidate("WPS Writer", "KWPS.Application", 17, "word"),
        "msoffice_word": BridgeCandidate("Microsoft Word", "Word.Application", 17, "word"),
    },
}


def _libreoffice_pdf_filter(category: str) -> str:
    if category == "spreadsheet":
        return "pdf:calc_pdf_Export"
    return "pdf:writer_pdf_Export"


def _configured_priority(context: ConverterContext, category: str) -> list[str]:
    default = _DEFAULT_PRIORITIES[category]
    configured = context.config.get(
        f"software.special_conversions.{category}_to_pdf",
        list(default),
    )
    if isinstance(configured, (list, tuple)):
        return [str(item) for item in configured if isinstance(item, str)]
    return list(default)


def _convert_with_configured_priority(
    context: ConverterContext,
    input_path: str,
    output_path: str,
    category: str,
    source_format: str,
) -> BridgeResult:
    return convert_with_backend_priority(
        input_path,
        output_path,
        source_format=source_format,
        backend_priority=_configured_priority(context, category),
        com_candidates=_COM_CANDIDATES[category],
        libreoffice_format=_libreoffice_pdf_filter(category),
        cancel=context.cancellation,
        failure_subject=f"Configured {category}→PDF backends",
    )


class OfficeToPdfConverter:
    """Convert Office-compatible inputs to PDF through the core office bridge."""

    def convert(self, context: ConverterContext) -> ConversionResult:
        task_id = context.request.request_id
        input_path = context.workspace.input_path
        output_path = context.workspace.create_artifact_path("primary", ".pdf")
        input_ref = context.request.input_refs[0]
        source_format = input_ref.format

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting Office → PDF conversion")

        bridge_result = _convert_with_configured_priority(
            context,
            input_path,
            output_path,
            input_ref.category,
            source_format,
        )
        if not bridge_result.success:
            message = bridge_result.message or "Office bridge conversion failed."
            if bridge_result.backend:
                message = f"{message} (backend: {bridge_result.backend})"
            return _error_result(task_id, message, "OFFICE2PDF-BRIDGE-FAILED")

        context.cancellation.check()
        context.progress.report_progress(100.0, "Office → PDF complete")
        source_stem = Path(input_path).stem
        result = _build_pdf_result(task_id, output_path, source_stem, source_format, context)
        result.metrics.extra["engine"] = "office_bridge"
        result.metrics.extra["backend"] = bridge_result.backend
        result.diagnostics.append(
            ConversionDiagnostic(
                level="info",
                message=bridge_result.message or "Office bridge conversion completed.",
                code="PDF-OFFICE-BRIDGE-OK",
            )
        )
        return result
