"""PluginRegistry — runtime component that registers and queries plugins.

The registry is the single source of truth for which plugins are available.
It wraps ``RouteRegistry`` from core with concrete plugin instances.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_core.formats.routes import RouteRegistry

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.protocols.converter import ConverterPlugin


class PluginRegistry:
    """Runtime registry of all loaded converter plugins.

    Responsibilities:
    - Accept plugin registrations (from discovery or manual injection).
    - Maintain a ``RouteRegistry`` for route → plugin lookups.
    - Provide plugin instances by id or by route match.

    This class is **not** a global singleton.  The runtime creates and
    owns the instance.

    Plugin package importing is owned by the distribution composition root
    (``docwen_bundle.runtime_factory``).  This registry only tracks concrete
    plugin instances that have already been loaded.
    """

    def __init__(self) -> None:
        self._plugins: dict[str, ConverterPlugin] = {}
        self._route_registry = RouteRegistry()

    # ── Registration ───────────────────────────────────────────────

    def register(self, plugin: ConverterPlugin) -> None:
        """Register a plugin and all its routes.

        If a plugin with the same ``plugin_id`` is already registered,
        it and its complete route set are replaced (last-register-wins).
        A replacement retains its registration order relative to other IDs.

        Args:
            plugin: A concrete ``ConverterPlugin`` instance.
        """
        manifest = plugin.manifest
        pid = manifest.plugin_id

        replacements = {**self._plugins, pid: plugin}
        self._replace_plugins(replacements)

    def unregister(self, plugin_id: str) -> None:
        """Remove a plugin and all its routes.

        No-op if *plugin_id* is not registered.
        """
        self._replace_plugins({pid: plugin for pid, plugin in self._plugins.items() if pid != plugin_id})

    def _replace_plugins(self, plugins: dict[str, ConverterPlugin]) -> None:
        """Build the next index before publishing either registry view."""
        routes = RouteRegistry()
        for pid, plugin in plugins.items():
            for route in plugin.manifest.routes:
                routes.register(route, pid)
        self._plugins = plugins
        self._route_registry = routes

    # ── Lookup ─────────────────────────────────────────────────────

    def get(self, plugin_id: str) -> ConverterPlugin | None:
        """Return a plugin by id, or ``None``."""
        return self._plugins.get(plugin_id)

    def find_plugin(
        self,
        source_format: str,
        target_format: str,
        action_name: str = "",
    ) -> ConverterPlugin | None:
        """Find a plugin that can handle the given route.

        Returns:
            The matching ``ConverterPlugin``, or ``None`` if no plugin
            is registered for this route.
        """
        entry = self._route_registry.find(source_format, target_format, action_name)
        if entry is None:
            return None
        return self._plugins.get(entry.plugin_id)

    def find_manifest(
        self,
        source_format: str,
        target_format: str,
        action_name: str = "",
    ) -> PluginManifest | None:
        """Find the manifest for a route.

        Returns:
            The ``PluginManifest`` of the matching plugin, or ``None``.
        """
        plugin = self.find_plugin(source_format, target_format, action_name)
        if plugin is None:
            return None
        return plugin.manifest

    # ── Introspection ──────────────────────────────────────────────

    @property
    def plugin_ids(self) -> list[str]:
        """Return all registered plugin ids."""
        return list(self._plugins.keys())

    @property
    def route_registry(self) -> RouteRegistry:
        """Return the underlying route registry (read-only intent)."""
        return self._route_registry

    def list_manifests(self) -> list[PluginManifest]:
        """Return manifests for all registered plugins."""
        return [p.manifest for p in self._plugins.values()]

    def __len__(self) -> int:
        return len(self._plugins)

    def __contains__(self, plugin_id: str) -> bool:
        return plugin_id in self._plugins
