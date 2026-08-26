"""docwen_plugin_image — image conversion plugin for DocWen.

Handles image → image, image → PDF, image → Markdown, and image merge to TIFF.

The plugin:
- Only depends on ``docwen_core``, Pillow, and img2pdf.
- Does NOT import runtime, application, gui, cli, bundle, old ``docwen``, or other plugins.
- Writes output to staging via ``WorkspaceHandle``.
- Returns ``ConversionResult`` with ``ArtifactManifest`` entries.
"""

from __future__ import annotations

from docwen_plugin_image.manifest import build_manifest
from docwen_plugin_image.plugin import ImagePlugin

__version__ = "0.1.0"
PLUGIN_CLASS = ImagePlugin
PLUGIN_MANIFEST = build_manifest()
__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "ImagePlugin", "__version__"]
