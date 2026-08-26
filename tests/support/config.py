"""Fake implementation of ReadOnlyConfigView protocol."""

from __future__ import annotations

from typing import Any


class FakeConfigView:
    """Fake config view with optional overrides.

    Usage::

        config = FakeConfigView({"some_key": "value"})
        config.get("some_key")  # → "value"
        config.get("missing", 42)  # → 42
    """

    def __init__(self, values: dict[str, Any] | None = None) -> None:
        self._values = dict(values) if values else {}

    def get(self, key: str, default: object = None) -> object:
        if key in self._values:
            return self._values[key]
        current: object = self._values
        for part in key.split("."):
            if not isinstance(current, dict) or part not in current:
                return default
            current = current[part]
        return current

    def get_plugin_config(self, plugin_id: str) -> dict[str, Any]:
        return self._values.get(plugin_id, {})
