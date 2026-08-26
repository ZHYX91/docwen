"""Fake implementation of PluginLogger protocol."""

from __future__ import annotations

from typing import Any


class FakePluginLogger:
    """Fake logger — records messages for test assertions.

    ``debug`` is intentionally discarded (not recorded) because debug-level
    messages are not asserted in tests.
    """

    def __init__(self) -> None:
        self.messages: list[dict[str, Any]] = []

    def debug(self, message: str, **extra: object) -> None:
        pass

    def info(self, message: str, **extra: object) -> None:
        self.messages.append({"level": "info", "message": message, **extra})

    def warning(self, message: str, **extra: object) -> None:
        self.messages.append({"level": "warning", "message": message, **extra})

    def error(self, message: str, **extra: object) -> None:
        self.messages.append({"level": "error", "message": message, **extra})
