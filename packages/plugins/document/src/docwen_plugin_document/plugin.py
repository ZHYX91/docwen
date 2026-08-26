"""DocumentPlugin — the entry point for docwen_plugin_document.

Implements the ``ConverterPlugin`` protocol from ``docwen_core``.
Handles DOCX → Markdown conversion and SmartConverter document-format
interconversion (requires external office software).

The plugin:
- Only depends on ``docwen_core`` and ``python-docx``.
- Does NOT import runtime, application, gui, cli, or other plugins.
- Writes output to staging via ``WorkspaceHandle``.
- Returns ``ConversionResult`` with ``ArtifactManifest`` entries.
- The runtime ``OutputFinalizer`` performs final placement.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_document.manifest import build_manifest
from docwen_plugin_document.to_document.converter import SmartDocConverter
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class DocumentPlugin:
    """Plugin focused on document family conversions (DOCX, DOC, ODT, RTF, WPS).

    Satisfies the ``ConverterPlugin`` protocol.

    Handles DOCX→Markdown and SmartConverter document-format
    interconversion (requires external office software).

    Optimizer-specific actions (gongwen, invoice_cn) are handled by
    dedicated plugins under packages/plugins/optimizers/.
    """

    plugin_id: str
    _manifest: PluginManifest | None
    _smart_converter: SmartDocConverter | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_document"
        self._manifest = None
        self._smart_converter = None

    @property
    def manifest(self) -> PluginManifest:
        """Return the plugin manifest (lazy-built)."""
        if self._manifest is None:
            self._manifest = build_manifest()
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        """Check if this plugin can handle the given route.

        Matches against all declared manifest routes so RouteResolver can
        resolve the document plugin before dispatch validates the exact
        implemented route.
        """
        for route in self.manifest.routes:
            if (
                route.source_format == source_format
                and route.target_format == target_format
                and route.action_name == action_name
            ):
                return True

        return False

    def convert(self, context: PluginExecutionContext) -> ConversionResult:
        """Dispatch to the appropriate converter based on source/target.

        Implemented routes:
          - docx/document → md  (DocxToMarkdownConverter)
          - SmartConverter document-format interconversion (SmartDocConverter, via external office bridge)
        """
        source = context.request.input_refs[0].format if context.request.input_refs else ""
        target = context.request.target_format

        # ── Implemented: DOCX / Document → Markdown ──────────────────
        if source in ("document", "docx") and target == "md":
            # Standard conversion owns mutable parsing state.  A fresh
            # instance keeps concurrent requests isolated without a lock.
            return DocxToMarkdownConverter().convert(context)

        # ── SmartConverter document-format interconversion (requires external office software) ──
        if source in SmartDocConverter.HANDLED_FORMATS and target in SmartDocConverter.HANDLED_FORMATS:
            if self._smart_converter is None:
                self._smart_converter = SmartDocConverter()
            return self._smart_converter.convert(context)

        # ── Fallback (should not normally be reached) ─────────────────
        return self._unsupported_route(context, f"{source}→{target}")

    @staticmethod
    def _unsupported_route(context: PluginExecutionContext, route_label: str) -> ConversionResult:
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        msg = f"{route_label} is not an executable document route."
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="DOCUMENT-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=msg,
                    code="DOCUMENT-UNSUPPORTED-ROUTE",
                )
            ],
        )
