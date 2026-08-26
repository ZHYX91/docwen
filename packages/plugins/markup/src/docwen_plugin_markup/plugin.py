"""MarkupPlugin — entry point for docwen_plugin_markup."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_markup.manifest import build_manifest
from docwen_plugin_markup.note_export.converter import EnexToMarkdownConverter
from docwen_plugin_markup.publication.converter import EpubToMarkdownConverter
from docwen_plugin_markup.web_archive.converter import HtmlToMarkdownConverter

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class MarkupPlugin:
    """Plugin for structured text / container formats to Markdown."""

    plugin_id: str
    _manifest: PluginManifest | None
    _html_converter: HtmlToMarkdownConverter | None
    _enex_converter: EnexToMarkdownConverter | None
    _epub_converter: EpubToMarkdownConverter | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_markup"
        self._manifest = None
        self._html_converter = None
        self._enex_converter = None
        self._epub_converter = None

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

        if source in ("html", "mhtml", "htm", "mht") and target == "md":
            if self._html_converter is None:
                self._html_converter = HtmlToMarkdownConverter()
            return self._html_converter.convert(context)

        if source == "enex" and target == "md":
            if self._enex_converter is None:
                self._enex_converter = EnexToMarkdownConverter()
            return self._enex_converter.convert(context)

        if source == "epub" and target == "md":
            if self._epub_converter is None:
                self._epub_converter = EpubToMarkdownConverter()
            return self._epub_converter.convert(context)

        return self._unsupported_route(context, f"{source}→{target}")

    @staticmethod
    def _unsupported_route(context: PluginExecutionContext, route_label: str) -> ConversionResult:
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        msg = f"{route_label} is not an executable markup route."
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="MARKUP-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[ConversionDiagnostic(level="error", message=msg, code="MARKUP-UNSUPPORTED-ROUTE")],
        )
