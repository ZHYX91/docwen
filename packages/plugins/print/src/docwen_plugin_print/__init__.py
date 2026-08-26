"""DocWen Print Plugin — paged output generation.

Handles generating fixed-layout output (PDF/OFD/XPS) from structured formats
(document, spreadsheet, etc.). Uses the office bridge infrastructure.

Depends on: docwen_core
Must NOT depend on: docwen_application, docwen_gui, docwen_cli,
                    docwen_bundle, or other plugin packages.
"""

from __future__ import annotations

from docwen_plugin_print.manifest import build_manifest
from docwen_plugin_print.plugin import PrintPlugin

__version__ = "0.1.0"
PLUGIN_CLASS = PrintPlugin
PLUGIN_MANIFEST = build_manifest()
__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "PrintPlugin", "__version__"]
