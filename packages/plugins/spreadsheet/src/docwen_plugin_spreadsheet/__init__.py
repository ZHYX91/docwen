"""docwen_plugin_spreadsheet — XLSX/CSV/TSV conversion plugin for DocWen.

This plugin handles:
- ROUTE-SHEET-001: spreadsheet → md (XLSX, CSV → Markdown)
- ROUTE-CSV-XLSX-001: csv → xlsx
- ROUTE-XLSX-CSV-001: xlsx → csv
- ROUTE-TSV-XLSX-001: tsv → xlsx
- ROUTE-XLSX-TSV-001: xlsx → tsv
- ROUTE-SHEETFMT-*: spreadsheet format interconversion (deferred — needs external office software)
- ACT-MERGE-TABLES: merge_tables

The plugin:
- Only depends on ``docwen_core``, ``openpyxl``, ``pandas``, and ``tabulate``.
- Does NOT import runtime, application, gui, cli, or other plugins.
- Writes output to staging via ``WorkspaceHandle``.
- Returns ``ConversionResult`` with ``ArtifactManifest`` entries.
- The runtime ``OutputFinalizer`` performs final placement.
"""

from __future__ import annotations

from docwen_plugin_spreadsheet.manifest import build_manifest
from docwen_plugin_spreadsheet.plugin import SpreadsheetPlugin

__version__ = "0.1.0"
PLUGIN_CLASS = SpreadsheetPlugin
PLUGIN_MANIFEST = build_manifest()
__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "SpreadsheetPlugin", "__version__"]
