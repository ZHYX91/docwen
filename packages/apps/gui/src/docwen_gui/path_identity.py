"""Shared path spelling for GUI lookups; does not resolve filesystem identity."""

from pathlib import Path


def normalize_path(file_path: str) -> str:
    """Normalize a path for internal lookup across widgets/view-models."""
    return str(Path(file_path)).replace("\\", "/")
