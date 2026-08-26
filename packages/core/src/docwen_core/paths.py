"""Path normalisation and safe-manipulation utilities for DocWen core.

All converters, preprocessors, and link resolvers should use these
functions instead of calling ``Path.resolve()`` or ``os.path.join``
directly — that ensures consistent behaviour across platforms and
prevents path-traversal vulnerabilities.

Audit mapping
  F-I2a-020  ``sanitize_filename``       → ``docwen_core.markdown_utils`` (resolved)
  F-I2a-021  ``sanitize_for_wiki_link``  → ``docwen_core.markdown_utils`` (resolved)
  (path normalisation / safe join utilities are new in this module)
"""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from pathlib import Path

IGNORED_INPUT_DIRECTORY_NAMES = frozenset(
    {
        ".git",
        "__pycache__",
        ".pytest_cache",
        "node_modules",
        ".venv",
        "venv",
        ".idea",
    }
)
WINDOWS_LEGACY_MAX_PATH = 260


def windows_utf16_units(path: str | os.PathLike[str]) -> int:
    """Return the Win32 path length unit used by the MAX_PATH boundary."""

    return len(os.fspath(path).encode("utf-16-le", errors="surrogatepass")) // 2


def _is_canonical_extended_path(path: str) -> bool:
    tail = path[4:]
    if len(tail) >= 3 and tail[0].isalpha() and tail[1:3] == ":\\":
        return True
    if not tail.upper().startswith("UNC\\"):
        return False
    components = tail[4:].split("\\")
    return len(components) >= 2 and all(component not in {"", ".", ".."} for component in components[:2])


def filesystem_path(
    path: str | os.PathLike[str],
    *,
    force_extended: bool = False,
) -> Path:
    """Return an absolute path suitable for an internal filesystem syscall.

    Public paths remain unchanged. On Windows, only the internal syscall
    operand receives a canonical extended prefix when it crosses MAX_PATH.
    Device and non-filesystem namespaces are rejected.
    """

    supplied = os.fspath(path)
    if sys.platform != "win32":
        return Path(os.path.abspath(supplied))
    supplied_namespace = supplied.replace("/", "\\")
    if supplied_namespace.startswith("\\\\.\\"):
        raise ValueError("Win32 device namespace paths are not supported")
    if supplied_namespace.startswith("\\\\?\\") and not _is_canonical_extended_path(supplied_namespace):
        raise ValueError("Win32 device namespace or non-filesystem extended namespace paths are not supported")

    raw = os.path.abspath(supplied)
    if raw.startswith("\\\\.\\"):
        raise ValueError("Win32 device namespace paths are not supported")
    if raw.startswith("\\\\?\\"):
        if not _is_canonical_extended_path(raw):
            raise ValueError("Win32 device namespace or non-filesystem extended namespace paths are not supported")
        return Path(raw)
    if not force_extended and windows_utf16_units(raw) < WINDOWS_LEGACY_MAX_PATH:
        return Path(raw)
    if raw.startswith("\\\\"):
        return Path(f"\\\\?\\UNC\\{raw[2:]}")
    return Path(f"\\\\?\\{raw}")


@dataclass(frozen=True, slots=True)
class DirectoryScanResult:
    """Files found by a bounded, non-link-following input-folder scan."""

    files: tuple[Path, ...]
    unreadable_paths: tuple[Path, ...] = ()
    truncated: bool = False


def scan_input_directory(
    directory: str | Path,
    *,
    limit: int | None = None,
) -> DirectoryScanResult:
    """Recursively enumerate ordinary files while pruning tool/cache trees.

    Directory-name matching is case-insensitive on every platform so a copied
    Windows project tree behaves identically on Linux. Symlink/reparse-point
    directories and symlink files are not followed into paths outside the
    selected root.
    """
    root = Path(directory)
    files: list[Path] = []
    unreadable: list[Path] = []
    truncated = False
    normalized_limit = None if limit is None else max(0, limit)

    def _record_error(error: OSError) -> None:
        unreadable.append(Path(error.filename) if error.filename else root)

    try:
        walker = os.walk(root, topdown=True, onerror=_record_error, followlinks=False)
        for current_root, dirnames, filenames in walker:
            current = Path(current_root)
            retained_directories: list[str] = []
            for dirname in sorted(dirnames, key=str.casefold):
                candidate = current / dirname
                if dirname.casefold() in IGNORED_INPUT_DIRECTORY_NAMES:
                    continue
                try:
                    if candidate.is_symlink() or candidate.is_junction():
                        continue
                except OSError:
                    unreadable.append(candidate)
                    continue
                retained_directories.append(dirname)
            dirnames[:] = retained_directories

            for filename in sorted(filenames, key=str.casefold):
                candidate = current / filename
                if normalized_limit is not None and len(files) >= normalized_limit:
                    truncated = True
                    break
                try:
                    if candidate.is_symlink() or not candidate.is_file():
                        continue
                except OSError:
                    unreadable.append(candidate)
                    continue
                files.append(candidate)
            if truncated:
                break
    except OSError as error:
        _record_error(error)

    files.sort(key=lambda path: str(path).casefold())
    return DirectoryScanResult(tuple(files), tuple(unreadable), truncated)


def normalize_path(path: str | Path) -> str:
    """Normalise *path* to an absolute, resolved, platform-consistent string.

    - Expands ``~`` and environment variables via :meth:`Path.expanduser`.
    - Resolves symlinks via :meth:`Path.resolve`.
    - On Windows the result is **lower-cased** so that case-insensitive
      filesystem comparisons are reliable.

    >>> normalize_path(Path.home() / "foo/bar")
    '/home/user/foo/bar'   # (platform-dependent)
    """
    resolved = Path(path).expanduser().resolve()
    if sys.platform == "win32":
        return str(resolved).lower()
    return str(resolved)


def safe_join_path(base: str | Path, *parts: str | Path) -> str:
    """Join *parts* onto *base* while preventing path-traversal attacks.

    The resulting path is resolved and checked to ensure it is still within
    *base*.  Raises :class:`ValueError` if the resolved path escapes the
    base directory.

    >>> safe_join_path("/home/user/docs", "images", "photo.png")
    '/home/user/docs/images/photo.png'
    >>> safe_join_path("/home/user/docs", "../../../etc/passwd")
    Traceback (most recent call last):
      …
    ValueError: Path traversal blocked: …
    """
    base_path = Path(base).expanduser().resolve(strict=False)
    target = base_path.joinpath(*parts).resolve(strict=False)

    # Path.is_relative_to was added in 3.9; use os.path.commonpath as fallback
    # but is_relative_to is cleaner and available on our minimum Python (3.11+).
    try:
        target.relative_to(base_path)
    except ValueError:
        raise ValueError(f"Path traversal blocked: {target} is outside {base_path}") from None

    return normalize_path(str(target))


def ensure_dir_exists(path: str | Path) -> str:
    """Ensure *path* exists as a directory, creating parents as needed.

    Returns the normalised directory path.

    >>> ensure_dir_exists("/tmp/docwen-test")
    '/tmp/docwen-test'
    """
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return normalize_path(str(p))


def input_stem(file_path: str | Path) -> str:
    """Return the stem (filename without extension) of *file_path*.

    Plugin code imports this helper when it needs shared filename semantics,
    instead of defining plugin-local wrapper functions in ``_common.py``.

    >>> input_stem("/a/b/c/财务报告_2024.pdf")
    '财务报告_2024'
    """
    return Path(file_path).stem


def input_name(file_path: str | Path) -> str:
    """Return the full filename (with extension) of *file_path*.

    >>> input_name("/a/b/c/财务报告_2024.pdf")
    '财务报告_2024.pdf'
    """
    return Path(file_path).name


__all__ = [
    "IGNORED_INPUT_DIRECTORY_NAMES",
    "WINDOWS_LEGACY_MAX_PATH",
    "DirectoryScanResult",
    "ensure_dir_exists",
    "filesystem_path",
    "input_name",
    "input_stem",
    "normalize_path",
    "safe_join_path",
    "scan_input_directory",
    "windows_utf16_units",
]
