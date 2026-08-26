"""Plugin registry support.

NOTE: This module (docwen_runtime.plugin_registry) is the runtime's internal
plugin REGISTRATION INFRASTRUCTURE.

It is NOT the same as packages/plugins/ — which is the directory where
individual plugin IMPLEMENTATIONS live (docwen_plugin_document, etc.).

The two are related but distinct:
- plugin_registry (here):  registers already-loaded plugin instances
- docwen_bundle:           imports default plugin packages for the distribution
- packages/plugins/:       contains the plugin packages themselves
"""

from docwen_runtime.plugin_registry.registry import PluginRegistry

__all__ = ["PluginRegistry"]
