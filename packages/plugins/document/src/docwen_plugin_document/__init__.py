"""DocWen Document Plugin — Document family ↔ Markdown conversion.

Depends on: docwen_core (and its own third-party libraries like python-docx)
Must NOT depend on: docwen_application, docwen_runtime, docwen_gui, docwen_cli,
                    docwen_bundle, or other plugin packages.
"""

__version__ = "0.1.0"

from docwen_plugin_document.manifest import build_manifest
from docwen_plugin_document.plugin import DocumentPlugin

PLUGIN_CLASS = DocumentPlugin
PLUGIN_MANIFEST = build_manifest()

__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "DocumentPlugin", "__version__"]
