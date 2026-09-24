"""Best-effort removal anchored to owned directories, without following links."""

from __future__ import annotations

import os
import stat
from contextlib import ExitStack
from pathlib import Path


def discard_file(root: Path, relative: Path) -> None:
    """Delete only a named file/link, then its empty parents; fail closed on I/O errors."""
    if not relative.parts or any(part in {".", ".."} for part in relative.parts):
        return
    try:
        if os.name == "nt":
            from .discard_windows import discard_windows

            discard_windows(root, relative)
        else:
            _discard_posix(root, relative)
    except OSError:
        # Cleanup must not replace the original conversion/validation error.
        return


def _discard_posix(root: Path, relative: Path) -> None:
    directory = getattr(os, "O_DIRECTORY", None)
    nofollow = getattr(os, "O_NOFOLLOW", None)
    if directory is None or nofollow is None:
        raise OSError("anchored directory cleanup is unavailable")
    flags = os.O_RDONLY | directory | nofollow
    with ExitStack() as stack:
        descriptor = os.open(root.anchor, flags)
        stack.callback(os.close, descriptor)
        # Anchor every lookup, including the staging root, to an open parent.
        for part in root.parts[1:]:
            descriptor = os.open(part, flags, dir_fd=descriptor)
            stack.callback(os.close, descriptor)
        parents: list[tuple[int, str, int]] = []
        for part in relative.parts[:-1]:
            child = os.open(part, flags, dir_fd=descriptor)
            stack.callback(os.close, child)
            parents.append((descriptor, part, child))
            descriptor = child
        metadata = os.stat(relative.name, dir_fd=descriptor, follow_symlinks=False)
        if not (stat.S_ISREG(metadata.st_mode) or stat.S_ISLNK(metadata.st_mode)):
            return
        # unlink never follows a leaf symlink. A swapped ancestor cannot redirect dir_fd.
        os.unlink(relative.name, dir_fd=descriptor)
        for parent, name, child in reversed(parents):
            current = os.stat(name, dir_fd=parent, follow_symlinks=False)
            opened = os.fstat(child)
            if (current.st_dev, current.st_ino) != (opened.st_dev, opened.st_ino):
                break
            os.rmdir(name, dir_fd=parent)
