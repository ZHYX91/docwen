"""docwen_plugin_proofread — rule-based text proofreading for DOCX and Markdown."""

from __future__ import annotations

from docwen_plugin_proofread.manifest import build_manifest
from docwen_plugin_proofread.plugin import ProofreadPlugin

PLUGIN_CLASS = ProofreadPlugin
PLUGIN_MANIFEST = build_manifest()

__all__ = ["PLUGIN_CLASS", "PLUGIN_MANIFEST", "ProofreadPlugin"]
