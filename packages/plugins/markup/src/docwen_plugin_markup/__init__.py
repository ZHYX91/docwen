"""DocWen Markup Plugin — HTML/ENEX/EPUB to Markdown conversion.

Depends on: docwen_core, markdownify
Must NOT depend on: docwen_application, docwen_runtime, docwen_gui, docwen_cli,
                    docwen_bundle, or other plugin packages.
"""

from __future__ import annotations

from docwen_plugin_markup.manifest import build_manifest
from docwen_plugin_markup.plugin import MarkupPlugin

__version__ = "0.1.0"
PLUGIN_CLASS = MarkupPlugin
PLUGIN_MANIFEST = build_manifest()
__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "MarkupPlugin", "__version__"]
