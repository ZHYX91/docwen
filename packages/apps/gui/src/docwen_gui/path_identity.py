"""Shared path spelling for GUI lookups; does not resolve filesystem identity."""

from pathlib import Path

from PySide6.QtCore import QDir


def normalize_path(file_path: str) -> str:
    """Normalize a path for internal lookup across widgets/view-models."""
    return str(Path(file_path)).replace("\\", "/")


def display_path(file_path: str) -> str:
    """Return a user-facing local path using the host platform separators."""
    if not file_path:
        return ""
    return QDir.toNativeSeparators(str(Path(file_path)))
