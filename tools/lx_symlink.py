"""Read WSL link metadata without asking Windows to follow the link."""

from __future__ import annotations

import struct
from pathlib import Path, PurePosixPath

LX_SYMLINK_TAG = 0xA000001D


def decode_target(data: bytes) -> str:
    if len(data) < 13:
        raise ValueError("invalid_lx_symlink_buffer")
    tag, length, _, version = struct.unpack_from("<IHHI", data)
    if tag != LX_SYMLINK_TAG or version != 2 or length != len(data) - 8:
        raise ValueError("invalid_lx_symlink_header")
    target = data[12:].decode("utf-8")
    # Guest absolute paths, drive/stream syntax and parent traversal need a
    # guest namespace resolver; never interpret them as Windows paths.
    if not target or any(c in target for c in ("\0", "\\", ":")):
        raise ValueError("unsafe_lx_symlink_target")
    if PurePosixPath(target).is_absolute() or ".." in PurePosixPath(target).parts:
        raise ValueError("unsupported_lx_symlink_target")
    return target


def read_target(path: Path) -> str:
    import win32file

    handle = win32file.CreateFile(
        str(path), 0, 7, None, 3, 0x00200000 | 0x02000000, None
    )  # OPEN_REPARSE_POINT | BACKUP_SEMANTICS; no target traversal
    try:
        data = win32file.DeviceIoControl(handle, 0x000900A8, None, 16384)
    finally:
        handle.Close()
    return decode_target(data)
