"""Route registry — helpers for matching conversion routes to plugins."""

from __future__ import annotations

from dataclasses import dataclass, field

from docwen_core.models.manifest import RouteSpec


@dataclass(slots=True)
class RouteEntry:
    """A registered conversion route linked to its owning plugin."""

    route: RouteSpec
    """The route specification."""

    plugin_id: str
    """The plugin that handles this route."""


@dataclass
class RouteRegistry:
    """A simple in-memory registry of conversion routes.

    Used by the runtime to resolve ``(source_format, target_format)``
    pairs to the correct plugin.  Not a global singleton — the runtime
    creates and owns the instance.

    ``RouteRegistry`` is a **runtime object** — it is built in memory
    during plugin discovery and does not need ``to_dict()/from_dict()``
    serialisation.  Individual ``RouteSpec`` and ``PluginManifest``
    objects are the serialisable units.
    """

    _entries: list[RouteEntry] = field(default_factory=list)

    def register(self, route: RouteSpec, plugin_id: str) -> None:
        """Register a route."""
        self._entries.append(RouteEntry(route=route, plugin_id=plugin_id))

    def find(
        self,
        source_format: str,
        target_format: str,
        action_name: str = "",
    ) -> RouteEntry | None:
        """Find a matching route entry.

        Precedence: exact action match > empty action match.
        """
        best: RouteEntry | None = None
        for entry in self._entries:
            r = entry.route
            if r.source_format != source_format or r.target_format != target_format:
                continue
            if r.action_name == action_name:
                return entry  # exact match — return immediately
            if not r.action_name:
                best = entry  # generic route — keep looking for exact match
        return best

    def list_routes(self) -> list[RouteEntry]:
        """Return all registered routes (copy)."""
        return list(self._entries)

    def list_plugin_routes(self, plugin_id: str) -> list[RouteEntry]:
        """Return all routes for a specific plugin."""
        return [e for e in self._entries if e.plugin_id == plugin_id]
