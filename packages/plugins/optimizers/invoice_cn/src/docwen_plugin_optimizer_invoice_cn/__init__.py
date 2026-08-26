"""docwen_plugin_optimizer_invoice_cn — Chinese invoice conversion plugin for DocWen.

Converts Chinese invoice files (PDF/OFD) to structured Markdown with
YAML frontmatter (20 metadata fields) and a detail-line table.

The plugin:
- Only depends on ``docwen_core`` and PyMuPDF.
- Does NOT import runtime, application, gui, cli, bundle, old ``docwen``, or other plugins.
- Writes output to staging via ``WorkspaceHandle``.
- Returns ``ConversionResult`` with ``ArtifactManifest`` entries.
- OCR-based image-to-invoice conversion is wired through the shared core OCR entry point.
"""

from __future__ import annotations

from docwen_plugin_optimizer_invoice_cn.manifest import build_manifest
from docwen_plugin_optimizer_invoice_cn.plugin import InvoicePlugin

__version__ = "0.1.0"
PLUGIN_CLASS = InvoicePlugin
PLUGIN_MANIFEST = build_manifest()
__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "InvoicePlugin", "__version__"]
