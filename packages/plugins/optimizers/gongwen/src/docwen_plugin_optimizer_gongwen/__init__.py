"""Gongwen (official Chinese document) recognition and rendering."""

from docwen_plugin_optimizer_gongwen.manifest import build_manifest
from docwen_plugin_optimizer_gongwen.plugin import GongwenOptimizerPlugin

PLUGIN_CLASS = GongwenOptimizerPlugin
PLUGIN_MANIFEST = build_manifest()
__version__ = "0.1.0"

__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "GongwenOptimizerPlugin", "__version__"]
