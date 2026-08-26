"""PresentationPlugin — entry point for docwen_plugin_presentation."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_presentation.manifest import build_manifest
from docwen_plugin_presentation.pptx_md.converter import PptxToMarkdownConverter
from docwen_plugin_presentation.pptx_md.ppt_converter import PptToMarkdownConverter

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class PresentationPlugin:
    """Plugin that converts PPTX/PPT presentations to Markdown."""

    plugin_id: str
    _manifest: PluginManifest | None
    _pptx_converter: PptxToMarkdownConverter | None
    _ppt_converter: PptToMarkdownConverter | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_presentation"
        self._manifest = None
        self._pptx_converter = None
        self._ppt_converter = None

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

        if source == "pptx" and target == "md":
            if self._pptx_converter is None:
                self._pptx_converter = PptxToMarkdownConverter()
            return self._pptx_converter.convert(context)

        if source == "ppt" and target == "md":
            if self._ppt_converter is None:
                self._ppt_converter = PptToMarkdownConverter()
            return self._ppt_converter.convert(context)

        return self._unsupported_route(context, f"{source}→{target}")

    @staticmethod
    def _unsupported_route(context: PluginExecutionContext, route_label: str) -> ConversionResult:
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        msg = f"{route_label} is not an executable presentation route."
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="PRESENTATION-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=msg,
                    code="PRESENTATION-UNSUPPORTED-ROUTE",
                )
            ],
        )
