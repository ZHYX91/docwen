"""MarkdownPlugin — entry point for docwen_plugin_markdown.

Implements the ``ConverterPlugin`` protocol from ``docwen_core``.
Handles Markdown → DOCX/WPS/XLSX/CSV conversion and heading numbering.

The plugin:
- Only depends on ``docwen_core``, ``python-docx``, ``openpyxl``, and ``mistune``.
- Does NOT import runtime, application, gui, cli, or other plugins.
- Writes output to staging via ``WorkspaceHandle``.
- Returns ``ConversionResult`` with ``ArtifactManifest`` entries.
- The runtime ``OutputFinalizer`` performs final placement.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docwen_core.models.file_inspection import FILE_INSPECTION_METADATA_KEY
from docwen_plugin_markdown.manifest import build_manifest
from docwen_plugin_markdown.numbering.converter import MdNumberingProcessor
from docwen_plugin_markdown.office_bridge.converter import MarkdownOfficeBridgeConverter
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_spreadsheet.converter import (
    MdToCsvConverter,
    MdToXlsxConverter,
)

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import DocumentStyleConverterContext, PluginExecutionContext


class MarkdownPlugin:
    """Plugin that converts Markdown to DOCX, XLSX, CSV and processes heading numbering.

    Satisfies the ``ConverterPlugin`` protocol.

    All 9 routes assigned by route_matrix.csv are declared in the manifest.
    Implemented routes:
      - markdown → docx  (MdToDocxConverter)
      - markdown → xlsx  (MdToXlsxConverter)
      - markdown → csv   (MdToCsvConverter)
      - markdown → doc/odt/rtf/wps/pdf/xls/ods (MarkdownOfficeBridgeConverter)
      - process_md_numbering (MdNumberingProcessor)
    """

    plugin_id: str
    _manifest: PluginManifest | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_markdown"
        self._manifest = None

    @property
    def manifest(self) -> PluginManifest:
        """Return the plugin manifest (lazy-built)."""
        if self._manifest is None:
            self._manifest = build_manifest()
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        """Check if this plugin can handle the given route.

        Matches against all declared manifest routes so RouteResolver can
        resolve the Markdown plugin before dispatch validates the exact
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
        """Dispatch to the appropriate converter based on source/target/action.

        Implemented routes:
          - markdown → docx  (MdToDocxConverter)
          - markdown → xlsx  (MdToXlsxConverter)
          - markdown → csv   (MdToCsvConverter)
          - markdown → doc/odt/rtf/wps/pdf/xls/ods (MarkdownOfficeBridgeConverter)
          - process_md_numbering (MdNumberingProcessor)
        """
        input_ref = context.request.input_refs[0] if context.request.input_refs else None
        source = input_ref.format if input_ref is not None else ""
        target = context.request.target_format
        action = context.request.action_name

        inspection = input_ref.metadata.get(FILE_INSPECTION_METADATA_KEY, {}) if input_ref is not None else {}
        declared_format = inspection.get("declared_format") if isinstance(inspection, dict) else None
        admitted_markdown_text = source == "txt" and declared_format == "markdown"
        source_format = "markdown" if source == "markdown" or admitted_markdown_text else source

        # ── Action: MD numbering ─────────────────────────────────
        if action == "process_md_numbering":
            return MdNumberingProcessor().convert(context)

        # ── Markdown → DOCX ──────────────────────────────────────
        if source_format == "markdown" and target == "docx":
            return MdToDocxConverter().convert(self._document_style_context(context))

        # ── Markdown → XLSX ──────────────────────────────────────
        if source_format == "markdown" and target == "xlsx":
            return MdToXlsxConverter().convert(context)

        # ── Markdown → CSV ───────────────────────────────────────
        if source_format == "markdown" and target == "csv":
            return MdToCsvConverter().convert(context)

        # ── Markdown → DOC/ODT/RTF/WPS/PDF/XLS/ODS via Office bridge ─────
        if source_format == "markdown" and target in (
            "doc",
            "odt",
            "rtf",
            "wps",
            "pdf",
            "xls",
            "ods",
        ):
            return MarkdownOfficeBridgeConverter().convert(context, target)

        # ── Fallback ─────────────────────────────────────────────
        return self._unsupported_route(context, f"{source}→{target}")

    @staticmethod
    def _document_style_context(context: PluginExecutionContext) -> DocumentStyleConverterContext:
        """Narrow a full plugin context only after the route requires DOCX styles."""

        if context.document_style_catalog is None:
            raise RuntimeError("DOCX route did not receive its request-owned document style catalog")
        return cast("DocumentStyleConverterContext", context)

    @staticmethod
    def _unsupported_route(context: PluginExecutionContext, route_label: str) -> ConversionResult:
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        msg = f"{route_label} is not an executable Markdown route."
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="MD-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=msg,
                    code="MD-UNSUPPORTED-ROUTE",
                )
            ],
        )
