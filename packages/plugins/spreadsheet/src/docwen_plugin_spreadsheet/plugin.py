"""SpreadsheetPlugin — the entry point for docwen_plugin_spreadsheet.

Implements the ``ConverterPlugin`` protocol from ``docwen_core``.
Handles XLSX/CSV/TSV → Markdown, CSV/TSV ↔ XLSX, format interconversion,
and table merging.

The plugin:
- Only depends on ``docwen_core``, ``openpyxl``, ``pandas``, and ``tabulate``.
- Does NOT import runtime, application, gui, cli, or other plugins.
- Writes output to staging via ``WorkspaceHandle``.
- Returns ``ConversionResult`` with ``ArtifactManifest`` entries.
- The runtime ``OutputFinalizer`` performs final placement.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_spreadsheet.csv_xlsx.converter import (
    CsvToXlsxConverter,
    TsvToXlsxConverter,
    XlsxToCsvConverter,
    XlsxToTsvConverter,
)
from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter
from docwen_plugin_spreadsheet.manifest import build_manifest
from docwen_plugin_spreadsheet.table_merger.converter import TableMergerConverter
from docwen_plugin_spreadsheet.to_markdown.converter import SpreadsheetToMarkdownConverter

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class SpreadsheetPlugin:
    """Plugin for spreadsheet format conversions.

    Satisfies the ``ConverterPlugin`` protocol.
    """

    plugin_id: str
    _manifest: PluginManifest | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_spreadsheet"
        self._manifest = None

    @property
    def manifest(self) -> PluginManifest:
        """Return the plugin manifest (lazy-built)."""
        if self._manifest is None:
            self._manifest = build_manifest()
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        """Check if this plugin can handle the given route.

        Derives supported routes directly from the manifest to guarantee
        consistency between ``can_handle()`` and declared ``RouteSpec``
        entries.
        """
        manifest = self.manifest
        for route in manifest.routes:
            if (
                route.source_format == source_format
                and route.target_format == target_format
                and route.action_name == action_name
            ):
                return True
        return False

    def convert(self, context: PluginExecutionContext) -> ConversionResult:
        """Run the appropriate conversion based on the request route.

        Dispatches to the correct converter based on source_format,
        target_format, and action_name.
        """
        source = context.request.input_refs[0].format if context.request.input_refs else ""
        target = context.request.target_format
        action = context.request.action_name

        # --- Action routes ---
        if action == "merge_tables":
            return TableMergerConverter().convert(context)

        # --- External-office bridge routes ---
        if source in ("xls", "ods", "et") and target == "md":
            return SmartSheetConverter().convert(context)

        # --- Core conversion routes: spreadsheet → md ---
        if target == "md":
            return SpreadsheetToMarkdownConverter().convert(context)

        if source == "csv" and target == "xlsx":
            return CsvToXlsxConverter().convert(context)

        if source == "xlsx" and target == "csv":
            return XlsxToCsvConverter().convert(context)

        if source == "tsv" and target == "xlsx":
            return TsvToXlsxConverter().convert(context)

        if source == "xlsx" and target == "tsv":
            return XlsxToTsvConverter().convert(context)

        # --- SmartSheetConverter routes (format interconversion) ---
        return SmartSheetConverter().convert(context)
