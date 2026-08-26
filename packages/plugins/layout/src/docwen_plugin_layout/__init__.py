"""docwen_plugin_layout — Layout conversion plugin for DocWen.

Handles:
- Fixed-layout PDF/OFD/XPS → Markdown via pymupdf4llm
- Fixed-layout PDF/OFD/XPS → PNG, JPG, TIF image rendering
- Fixed-layout PDF/OFD/XPS → DOCX, DOC, ODT, RTF (requires external Office)
- Fixed-layout PDF/OFD/XPS → PDF (normalize / pass-through)

The plugin:
- Only depends on ``docwen_core``, PyMuPDF, and optional easyofd.
- Does NOT import runtime, application, gui, cli, bundle, old ``docwen``, or other plugins.
- Writes output to staging via ``WorkspaceHandle``.
- Returns ``ConversionResult`` with ``ArtifactManifest`` entries.
"""

from __future__ import annotations

from docwen_plugin_layout.manifest import build_manifest
from docwen_plugin_layout.plugin import LayoutPlugin

__version__ = "0.1.0"
PLUGIN_CLASS = LayoutPlugin
PLUGIN_MANIFEST = build_manifest()
__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "LayoutPlugin", "__version__"]
