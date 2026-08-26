"""ConverterPlugin protocol — the interface every plugin converter must implement."""

from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class ConverterPlugin(Protocol):
    """Protocol that every plugin converter must satisfy.

    A plugin is a callable that accepts a ``PluginExecutionContext``
    and returns a ``ConversionResult``.  It also exposes a ``manifest``
    property so the runtime can discover its capabilities.

    Plugins MUST NOT:
    - Write directly to the final output directory.
    - Hold mutable global configuration.
    - Import ``docwen_runtime``, ``docwen_application``, ``docwen_gui``, or ``docwen_cli``.
    """

    @property
    def manifest(self) -> PluginManifest:
        """Metadata describing this plugin and its routes."""
        ...

    def convert(self, context: PluginExecutionContext) -> ConversionResult:
        """Execute a conversion.

        The plugin receives everything it needs via *context*:
        - The input file reference
        - A workspace handle for staging writes
        - A read-only config view
        - A progress sink for reporting
        - A cancellation token for cooperative cancellation

        Returns a ``ConversionResult`` with staging artifacts and diagnostics.
        """
        ...

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        """Return ``True`` if this plugin can convert *source_format* → *target_format*.

        If *action_name* is non-empty the plugin should also check that the
        action matches (e.g. ``"validate"``, ``"merge"``).
        """
        ...
