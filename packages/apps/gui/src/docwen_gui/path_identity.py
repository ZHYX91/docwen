"""Host-aware path spelling and lookup keys, without resolving symlinks."""

import os
from pathlib import Path

from PySide6.QtCore import QDir


def normalize_path(file_path: str) -> str:
    """Keep file-name spelling while normalizing the host's separators.

    A backslash is a valid POSIX filename character, not a separator.
    This value may be displayed or passed to filesystem APIs; use ``path_key``
    for identity comparisons instead.
    """
    return str(Path(file_path)).replace(os.sep, "/") if file_path else ""


def path_key(file_path: str | os.PathLike[str]) -> str:
    """Return the shared host-native lookup key, preserving POSIX backslashes.

    Windows case/separator aliases compare equally; relative spellings are
    anchored to the current working directory. Symlinks are not resolved.
    Keys must not replace the original source path in a FileRef or UI label.
    """
    raw = os.fspath(file_path)
    return os.path.normcase(os.path.abspath(raw)) if raw else ""


def display_path(file_path: str) -> str:
    """Return a user-facing local path using the host platform separators."""
    if not file_path:
        return ""
    return QDir.toNativeSeparators(str(Path(file_path)))
