"""Recycle validated files or directories and prove a recoverable payload exists."""

from __future__ import annotations

import os
import struct
import time
from pathlib import Path


def _perform_recycle(path: Path) -> None:
    import pythoncom
    from win32com.shell import shell

    pythoncom.CoInitialize()
    try:
        operation = pythoncom.CoCreateInstance(
            shell.CLSID_FileOperation, None, pythoncom.CLSCTX_ALL, shell.IID_IFileOperation
        )
        # RECYCLEONDELETE, EARLYFAILURE, ALLOWUNDO, NOERRORUI, NOCONFIRMATION, SILENT.
        # RECYCLEONDELETE fails when recycling is unavailable; no permanent fallback.
        operation.SetOperationFlags(0x80000 | 0x100000 | 0x40 | 0x400 | 0x10 | 0x4)
        operation.DeleteItem(shell.SHCreateItemFromParsingName(str(path), None, shell.IID_IShellItem), None)
        operation.PerformOperations()
        if operation.GetAnyOperationsAborted():
            raise OSError(f"recycle_aborted:{path}")
    finally:
        pythoncom.CoUninitialize()


def _recovery_entries(path: Path) -> list[dict[str, str]]:
    matches = []
    root = Path(path.anchor) / "$Recycle.Bin"
    if not root.exists():
        return []
    for owner in root.iterdir():
        try:
            for metadata in owner.glob("$I*"):
                try:
                    content = metadata.read_bytes()
                    version = struct.unpack_from("<Q", content)[0]
                    if version not in {1, 2}:
                        continue
                    original = content[28 if version == 2 else 24 :].decode("utf-16-le").split("\0")[0]
                    payload = metadata.with_name("$R" + metadata.name[2:])
                    if os.path.normcase(original) == os.path.normcase(str(path)) and payload.exists():
                        matches.append({"metadata": str(metadata), "payload": str(payload)})
                except (OSError, ValueError, UnicodeError, struct.error):
                    continue
        except OSError:
            continue
    return matches


def recycle_path(path: Path) -> list[dict[str, str]]:
    """Return new recovery entries, or fail without attempting permanent deletion."""
    if os.name != "nt":
        raise OSError("recycle_requires_windows")
    if not path.is_absolute() or path == Path(path.anchor):
        raise ValueError("recycle_requires_absolute_nonroot_target")
    if path.is_symlink() or getattr(path.lstat(), "st_file_attributes", 0) & 0x400:
        raise ValueError("recycle_target_is_reparse_point")
    path = path.resolve(strict=True)
    previous = {entry["metadata"] for entry in _recovery_entries(path)}
    _perform_recycle(path)
    if path.exists():
        raise OSError(f"recycle_target_remains:{path}")
    # Shell return can precede visibility of the metadata entry.
    for _ in range(30):
        entries = [entry for entry in _recovery_entries(path) if entry["metadata"] not in previous]
        if entries:
            return entries
        time.sleep(0.1)
    raise OSError(f"recycle_recovery_unverified:{path}")


def recycle_directory(path: Path) -> list[dict[str, str]]:
    """Compatibility entry point for saved directory cleanup plans."""
    return recycle_path(path)
