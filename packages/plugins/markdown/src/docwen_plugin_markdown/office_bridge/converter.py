"""Bridge Markdown exports through generated Office intermediates."""

from __future__ import annotations

import time
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.result import ConversionDiagnostic, ConversionErrorInfo, ConversionMetrics, ConversionResult
from docwen_core.office_bridge import BridgeCandidate, convert_with_backend_priority
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_spreadsheet.converter import MdToXlsxConverter

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import (
        ConverterContext,
        DocumentStyleConverterContext,
        PluginExecutionContext,
    )


_DOCUMENT_TARGETS = {"doc", "odt", "rtf", "wps", "pdf"}
_SPREADSHEET_TARGETS = {"xls", "ods"}

_MEDIA_TYPES = {
    "doc": "application/msword",
    "odt": "application/vnd.oasis.opendocument.text",
    "rtf": "application/rtf",
    "wps": "application/msword",
    "pdf": "application/pdf",
    "xls": "application/vnd.ms-excel",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
}

_WORD_SAVE_FORMATS = {
    "doc": 0,
    "odt": 23,
    "rtf": 6,
    "wps": 0,
    "pdf": 17,
}

_EXCEL_SAVE_FORMATS = {
    "xls": 56,
    "ods": 60,
}

_LIBREOFFICE_FORMATS = {
    "doc": "doc",
    "odt": "odt",
    "rtf": "rtf",
    "wps": "doc",
    "pdf": "pdf",
    "xls": "xls",
    "ods": "ods",
}

_DEFAULT_PRIORITIES = {
    "word_processors": ("wps_writer", "msoffice_word", "libreoffice"),
    "spreadsheet_processors": ("wps_spreadsheets", "msoffice_excel", "libreoffice"),
    "odt": ("msoffice_word", "libreoffice"),
    "ods": ("msoffice_excel", "libreoffice"),
    "document_to_pdf": ("wps_writer", "msoffice_word", "libreoffice"),
}

_PRIORITY_KEYS = {
    "word_processors": "software.default_priority.word_processors",
    "spreadsheet_processors": "software.default_priority.spreadsheet_processors",
    "odt": "software.special_conversions.odt",
    "ods": "software.special_conversions.ods",
    "document_to_pdf": "software.special_conversions.document_to_pdf",
}


class MarkdownOfficeBridgeConverter:
    """Convert Markdown to legacy Office formats via DOCX/XLSX intermediates."""

    def convert(self, context: PluginExecutionContext, target_format: str) -> ConversionResult:
        if target_format in _DOCUMENT_TARGETS:
            return self._convert_with_intermediate(
                context,
                target_format,
                lambda full_context: MdToDocxConverter().convert(self._document_style_context(full_context)),
                "docx",
            )
        if target_format in _SPREADSHEET_TARGETS:
            return self._convert_with_intermediate(
                context,
                target_format,
                MdToXlsxConverter().convert,
                "xlsx",
            )
        return self._unsupported_result(context, target_format)

    @staticmethod
    def _document_style_context(context: PluginExecutionContext) -> DocumentStyleConverterContext:
        from typing import cast

        if context.document_style_catalog is None:
            raise RuntimeError("DOCX bridge route did not receive its request-owned document style catalog")
        return cast("DocumentStyleConverterContext", context)

    def _convert_with_intermediate(
        self,
        context: PluginExecutionContext,
        target_format: str,
        intermediate_converter: Callable[[PluginExecutionContext], ConversionResult],
        intermediate_format: str,
    ) -> ConversionResult:
        t_start = time.monotonic()
        task_id = context.request.request_id
        try:
            intermediate_result = intermediate_converter(context)
            if not intermediate_result.success or not intermediate_result.artifacts:
                return intermediate_result

            intermediate_artifact = intermediate_result.artifacts[0]
            intermediate_path = Path(intermediate_artifact.staging_path)
            input_stem = Path(context.workspace.input_path).stem
            output_path = Path(context.workspace.create_artifact_path(input_stem, f".{target_format}"))
            bridge_output_path = output_path.with_suffix(".doc") if target_format == "wps" else output_path

            priority_category = self._priority_category(target_format)
            bridge_result = convert_with_backend_priority(
                str(intermediate_path),
                str(bridge_output_path),
                source_format=intermediate_format,
                backend_priority=self._configured_priority(context, priority_category),
                com_candidates=self._candidates_for(target_format),
                libreoffice_format=_LIBREOFFICE_FORMATS[target_format],
                cancel=context.cancellation,
                failure_subject=f"Configured Markdown→{target_format.upper()} backends",
            )
            if not bridge_result.success:
                return self._bridge_error_result(context, target_format, bridge_result.message, t_start)
            if target_format == "wps" and bridge_output_path.exists():
                bridge_output_path.replace(output_path)

            output_bytes = output_path.stat().st_size if output_path.exists() else 0
            artifact = ArtifactManifest(
                artifact_id=f"{task_id}-{target_format}",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=str(output_path),
                suggested_name=f"{input_stem}.{target_format}",
                media_type=_MEDIA_TYPES[target_format],
                is_primary=True,
                metadata={
                    "source_format": "markdown",
                    "target_format": target_format,
                    "intermediate_format": intermediate_format,
                    "engine": "office_bridge",
                    "backend": bridge_result.backend,
                },
            )
            context.workspace.add_artifact(artifact)

            elapsed_ms = (time.monotonic() - t_start) * 1000.0
            input_bytes = Path(context.workspace.input_path).stat().st_size
            return ConversionResult(
                task_id=task_id,
                success=True,
                artifacts=[artifact],
                metrics=ConversionMetrics(
                    duration_ms=elapsed_ms,
                    input_bytes=input_bytes,
                    output_bytes=output_bytes,
                    extra={"engine": "office_bridge", "backend": bridge_result.backend},
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="info",
                        message=f"Markdown→{target_format.upper()} conversion successful via Office bridge",
                        code="MD-OFFICE-BRIDGE-OK",
                    )
                ],
            )
        except Exception as exc:
            return self._bridge_error_result(context, target_format, str(exc), t_start)

    @staticmethod
    def _priority_category(target_format: str) -> str:
        if target_format == "odt":
            return "odt"
        if target_format == "ods":
            return "ods"
        if target_format == "pdf":
            return "document_to_pdf"
        if target_format in _DOCUMENT_TARGETS:
            return "word_processors"
        return "spreadsheet_processors"

    @staticmethod
    def _configured_priority(context: ConverterContext, category: str) -> list[str]:
        default = _DEFAULT_PRIORITIES[category]
        configured = context.config.get(_PRIORITY_KEYS[category], list(default))
        if isinstance(configured, (list, tuple)):
            return [str(item) for item in configured if isinstance(item, str)]
        return list(default)

    def _candidates_for(self, target_format: str) -> dict[str, BridgeCandidate]:
        if target_format in _DOCUMENT_TARGETS:
            save_format = _WORD_SAVE_FORMATS[target_format]
            candidates = {
                "wps_writer": BridgeCandidate("WPS Writer", "KWps.Application", save_format, "word"),
                "msoffice_word": BridgeCandidate("Microsoft Word", "Word.Application", save_format, "word"),
            }
            if target_format == "odt":
                candidates.pop("wps_writer")
            return candidates
        save_format = _EXCEL_SAVE_FORMATS[target_format]
        candidates = {
            "wps_spreadsheets": BridgeCandidate("WPS Spreadsheet", "KET.Application", save_format, "excel"),
            "msoffice_excel": BridgeCandidate("Microsoft Excel", "Excel.Application", save_format, "excel"),
        }
        if target_format == "ods":
            candidates.pop("wps_spreadsheets")
        return candidates

    def _unsupported_result(self, context: ConverterContext, target_format: str) -> ConversionResult:
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=f"Markdown Office bridge does not support target format: {target_format}",
                diagnostic_code="MD-OFFICE-BRIDGE-UNSUPPORTED",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=f"Unsupported Markdown Office bridge target: {target_format}",
                    code="MD-OFFICE-BRIDGE-UNSUPPORTED",
                )
            ],
        )

    def _bridge_error_result(
        self,
        context: ConverterContext,
        target_format: str,
        message: str,
        t_start: float,
    ) -> ConversionResult:
        elapsed_ms = (time.monotonic() - t_start) * 1000.0
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="office_bridge_failed",
                message=f"Markdown→{target_format.upper()} Office bridge conversion failed: {message}",
                diagnostic_code="MD-OFFICE-BRIDGE-FAILED",
            ),
            metrics=ConversionMetrics(duration_ms=elapsed_ms),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=f"Markdown→{target_format.upper()} Office bridge conversion failed: {message}",
                    code="MD-OFFICE-BRIDGE-ERROR",
                )
            ],
        )
