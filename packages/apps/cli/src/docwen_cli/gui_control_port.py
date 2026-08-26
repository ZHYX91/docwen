"""Bundle-neutral port for DocWen GUI control operations."""

from __future__ import annotations

from typing import Any, Protocol

GUI_SETTINGS_SECTIONS = ("proofread",)


class GuiControlError(Exception):
    """Typed failure returned by a GUI control port."""

    def __init__(self, code: str, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.details = details or {}


class GuiControlPort(Protocol):
    """Operations required by the CLI without owning a runtime transport."""

    def status(self, *, timeout: float) -> dict[str, Any]: ...

    def activate(self, *, timeout: float) -> dict[str, Any]: ...

    def open(self, file_path: str | None, *, timeout: float) -> dict[str, Any]: ...

    def open_settings(self, section: str, *, timeout: float) -> dict[str, Any]: ...


__all__ = ["GUI_SETTINGS_SECTIONS", "GuiControlError", "GuiControlPort"]
