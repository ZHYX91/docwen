"""RouteResolver — resolves a conversion request to a specific plugin and route."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_core.models.manifest import RouteSpec

if TYPE_CHECKING:
    from docwen_core.models.file_ref import FileRef
    from docwen_runtime.plugin_registry.registry import PluginRegistry


class RouteResolutionError(Exception):
    """Raised when no plugin can handle the requested conversion."""

    def __init__(self, source_format: str, target_format: str, action_name: str = "") -> None:
        self.source_format = source_format
        self.target_format = target_format
        self.action_name = action_name
        msg = f"No plugin found for {source_format!r} → {target_format!r}"
        if action_name:
            msg += f" (action={action_name!r})"
        super().__init__(msg)


class RouteResolver:
    """Resolves ``(source_format, target_format, action_name)`` to a plugin.

    Wraps the ``PluginRegistry`` and provides resolution logic with
    clear error reporting.  This is a thin orchestration layer —
    the actual route matching is delegated to ``RouteRegistry.find()``.

    The concrete content-derived ``FileRef.format`` is tried first.  A route
    may intentionally be registered for the broader workflow category (for
    example ``markdown`` or ``image``), in which case ``FileRef.category`` is
    the only fallback.  Paths and suffixes never participate in routing.
    """

    def __init__(self, plugin_registry: PluginRegistry) -> None:
        self._registry = plugin_registry

    def resolve(
        self,
        input_ref: FileRef,
        target_format: str,
        action_name: str = "",
    ) -> tuple[str, RouteSpec]:
        """Resolve a conversion route.

        Args:
            input_ref: The input file reference (provides source format).
            target_format: Desired output format.
            action_name: Optional named action (validate, merge, etc.).

        Returns:
            A ``(plugin_id, RouteSpec)`` tuple.

        Raises:
            RouteResolutionError: If no plugin handles this route.
        """
        source_format = str(input_ref.format or "")
        source_category = str(input_ref.category or "")
        candidates = tuple(
            candidate for candidate in (source_format, source_category) if candidate and candidate != "unknown"
        )
        entry = None
        for candidate in dict.fromkeys(candidates):
            entry = self._registry.route_registry.find(candidate, target_format, action_name)
            if entry is not None:
                break
        if entry is None:
            raise RouteResolutionError(source_format or "unknown", target_format, action_name)

        return entry.plugin_id, entry.route

    def resolve_plugin_id(
        self,
        input_ref: FileRef,
        target_format: str,
        action_name: str = "",
    ) -> str:
        """Resolve to a plugin id only (convenience)."""
        plugin_id, _ = self.resolve(input_ref, target_format, action_name)
        return plugin_id
