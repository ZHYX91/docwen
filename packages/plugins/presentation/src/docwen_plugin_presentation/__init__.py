"""DocWen Presentation Plugin — PPTX/PPT to Markdown conversion.

Depends on: docwen_core, python-pptx
Must NOT depend on: docwen_application, docwen_runtime, docwen_gui, docwen_cli,
                    docwen_bundle, or other plugin packages.
"""

from __future__ import annotations

from docwen_plugin_presentation.manifest import build_manifest
from docwen_plugin_presentation.plugin import PresentationPlugin

__version__ = "0.1.0"
PLUGIN_CLASS = PresentationPlugin
PLUGIN_MANIFEST = build_manifest()
__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "PresentationPlugin", "__version__"]
