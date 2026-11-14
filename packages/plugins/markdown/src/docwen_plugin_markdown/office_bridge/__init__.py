"""Lazy public converter exports; importing a utility never loads its engine."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .converter import MarkdownOfficeBridgeConverter

__all__ = ["MarkdownOfficeBridgeConverter"]


def __getattr__(name: str) -> Any:
    if name not in __all__:
        raise AttributeError(name)
    from . import converter

    value = getattr(converter, name)
    globals()[name] = value
    return value
