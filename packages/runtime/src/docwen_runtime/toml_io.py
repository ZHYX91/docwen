"""Shared TOML I/O primitives for runtime, GUI, and plugins."""

from __future__ import annotations

import contextlib
import ctypes
import os
import stat
import tempfile
import tomllib
from pathlib import Path
from typing import Any

import tomlkit


def _sync_windows_directory(directory: Path) -> None:
    """Flush directory metadata using the Windows backup-semantics handle."""
    from ctypes import wintypes

    ctypes_api = vars(ctypes)
    win_dll = ctypes_api["WinDLL"]
    win_error = ctypes_api["WinError"]
    get_last_error = ctypes_api["get_last_error"]
    kernel32 = win_dll("kernel32", use_last_error=True)
    create_file = kernel32.CreateFileW
    create_file.argtypes = (
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    )
    create_file.restype = wintypes.HANDLE
    flush_file_buffers = kernel32.FlushFileBuffers
    flush_file_buffers.argtypes = (wintypes.HANDLE,)
    flush_file_buffers.restype = wintypes.BOOL
    close_handle = kernel32.CloseHandle
    close_handle.argtypes = (wintypes.HANDLE,)
    close_handle.restype = wintypes.BOOL

    generic_write = 0x40000000
    share_read_write_delete = 0x00000007
    open_existing = 3
    file_flag_backup_semantics = 0x02000000
    invalid_handle_value = ctypes.c_void_p(-1).value
    handle = create_file(
        str(directory),
        generic_write,
        share_read_write_delete,
        None,
        open_existing,
        file_flag_backup_semantics,
        None,
    )
    if handle == invalid_handle_value:
        raise win_error(get_last_error())
    try:
        if not flush_file_buffers(handle):
            raise win_error(get_last_error())
    finally:
        close_handle(handle)


def _sync_directory(path: str | Path) -> None:
    """Durably order one local directory's entry metadata."""
    directory = Path(path)
    if os.name == "nt":
        _sync_windows_directory(directory)
        return
    flags = os.O_RDONLY | getattr(os, "O_DIRECTORY", 0)
    descriptor = os.open(directory, flags)
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


def atomic_write_bytes(path: str | Path, content: bytes) -> None:
    """Atomically replace *path* with *content* from a same-directory temp file.

    The temporary file is flushed and synced before ``os.replace``.  Keeping
    it beside the destination prevents a handled temp-write failure from
    truncating the target and avoids cross-filesystem replacement; it does not
    promise multi-file crash atomicity or power-loss durability.
    """
    file_path = Path(path)
    if file_path.is_symlink():
        file_path = file_path.resolve(strict=False)
    file_path.parent.mkdir(parents=True, exist_ok=True)
    existing_mode: int | None = None
    with contextlib.suppress(FileNotFoundError):
        existing_mode = stat.S_IMODE(file_path.stat().st_mode)

    descriptor, temp_name = tempfile.mkstemp(
        dir=file_path.parent,
        prefix=f".{file_path.name}.",
        suffix=".tmp",
    )
    temp_path = Path(temp_name)
    descriptor_open = True
    try:
        with os.fdopen(descriptor, "wb") as stream:
            descriptor_open = False
            stream.write(content)
            stream.flush()
            if existing_mode is not None:
                os.chmod(temp_path, existing_mode)
            os.fsync(stream.fileno())
        os.replace(temp_path, file_path)
        _sync_directory(file_path.parent)
    finally:
        if descriptor_open:
            os.close(descriptor)
        with contextlib.suppress(FileNotFoundError):
            temp_path.unlink()


def durable_unlink(path: str | Path, *, missing_ok: bool = False) -> bool:
    """Unlink one entry and durably order its parent directory.

    Returns ``True`` only when an entry was removed.  Broken symlinks count as
    entries.  ``missing_ok`` controls the ordinary missing-path case.
    """
    file_path = Path(path)
    try:
        file_path.unlink()
    except FileNotFoundError:
        if missing_ok:
            return False
        raise
    _sync_directory(file_path.parent)
    return True


def atomic_write_text(path: str | Path, content: str) -> None:
    """UTF-8 encode *content* and atomically replace *path*."""
    atomic_write_bytes(path, content.encode("utf-8"))


def read_toml_file(path: str | Path) -> dict[str, Any]:
    """Read TOML into a plain dictionary without edit-preservation overhead."""
    with Path(path).open("rb") as stream:
        return tomllib.load(stream)


def write_toml_file(path: str | Path, data: dict[str, Any]) -> None:
    """Serialize TOML data with a same-directory staged replacement."""
    atomic_write_text(path, tomlkit.dumps(data))


def new_toml_document() -> Any:
    """Create an empty mutable TOML document."""
    return tomlkit.document()


def load_toml_document(path: str | Path) -> Any:
    """Read a TOML file as a **mutable** tomlkit document (preserving comments and formatting).

    Use this only when the caller will mutate the document and write it back
    via :func:`save_toml_document`.
    """
    return tomlkit.parse(Path(path).read_text(encoding="utf-8"))


def save_toml_document(path: str | Path, doc: Any) -> None:
    """Write a mutable TOML document via same-directory staged replacement."""
    atomic_write_text(path, doc.as_string())


def update_toml_document_sections(doc: Any, updates: dict[str, Any]) -> None:
    """Merge top-level *updates* sections into a tomlkit document *doc*.

    For each top-level key in *updates*:
    - If *doc* already has that key, **replace the entire section** with
      the value from *updates* (not a per-key deep merge).  This matches
      the real intent behind the old ``tomlkit.parse(tomlkit.dumps(data))``
      anti-pattern: take a section the editor rebuilt in memory and write
      it back wholesale.
    - If *doc* does not have the key, add the new section.

    **Comment preservation scope (important constraint):** Only sections
    **not** appearing in *updates* keep their comments, key order, and
    inline comments.  All comments inside a replaced section (section-
    header comments, per-value inline comments) are **lost** because the
    implementation uses ``parse(dumps({key: value}))`` internally, and
    tomlkit's ``dumps`` does not emit comments.

    This means: if a caller needs to preserve per-value inline comments
    inside a section (e.g. user remarks on each entry in a proofread
    dictionary), **do not use this function** -- use
    ``ConfigLoader.update_file_document`` to obtain the tomlkit document
    and manipulate it at the fine-grained level.  This function is meant
    for "wholesale section rebuild" scenarios such as numbering schemes
    or clean rules -- configuration that has no per-value comments.

    *updates* values are plain dicts (from the editor's in-memory model).
    They are converted to tomlkit-typed sub-documents via
    ``tomlkit.parse(tomlkit.dumps(...))`` before assignment so that
    compound types (arrays-of-tables, etc.) serialize correctly.
    """
    for key, value in updates.items():
        if not isinstance(value, dict):
            doc[key] = value
            continue
        doc[key] = tomlkit.parse(tomlkit.dumps({key: value}))[key]
