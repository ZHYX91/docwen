"""PrintPlugin — entry point for docwen_plugin_print.

Generates fixed-layout PDF output from structured input formats.
Office bridge-backed routes (document/spreadsheet→pdf) are wired through
OfficeToPdfConverter in paged_output.converter.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_print.manifest import build_manifest
from docwen_plugin_print.paged_output.converter import OfficeToPdfConverter

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class PrintPlugin:
    """Plugin for generating paged output from structured formats."""

    plugin_id: str
    _manifest: PluginManifest | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_print"
        self._manifest = None

    @property
    def manifest(self) -> PluginManifest:
        if self._manifest is None:
            self._manifest = build_manifest()
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        for route in self.manifest.routes:
            if (
                route.source_format == source_format
                and route.target_format == target_format
                and route.action_name == action_name
            ):
                return True
        return False

    def convert(self, context: PluginExecutionContext) -> ConversionResult:
        source = context.request.input_refs[0].format if context.request.input_refs else ""
        target = context.request.target_format

        # Document → PDF (Office bridge)
        if source in ("document", "docx", "doc", "rtf", "odt", "wps") and target == "pdf":
            return OfficeToPdfConverter().convert(context)

        # Spreadsheet → PDF (Office bridge)
        if source in ("spreadsheet", "xlsx", "xls", "et", "ods", "csv") and target == "pdf":
            return OfficeToPdfConverter().convert(context)

        return self._unsupported_route(context, f"{source}→{target}")

    @staticmethod
    def _unsupported_route(context: PluginExecutionContext, route_label: str) -> ConversionResult:
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        msg = f"{route_label} is not an executable print route."
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="PRINT-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[ConversionDiagnostic(level="error", message=msg, code="PRINT-UNSUPPORTED-ROUTE")],
        )
