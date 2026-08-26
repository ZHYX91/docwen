"""DocWen Markdown Plugin — Markdown conversion engine."""

from docwen_plugin_markdown.manifest import build_manifest
from docwen_plugin_markdown.plugin import MarkdownPlugin

PLUGIN_CLASS = MarkdownPlugin
PLUGIN_MANIFEST = build_manifest()

__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "MarkdownPlugin"]
