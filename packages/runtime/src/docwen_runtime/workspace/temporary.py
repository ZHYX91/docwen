"""Bounded, cancellation-aware private directories.

Windows can report PermissionError on every mkdir while a parent still appears
writable. tempfile.mkdtemp retries that condition almost indefinitely. Retry
only an occupied candidate; do not retry a persistent access denial.
"""

from __future__ import annotations

import os
import secrets
import shutil
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import CancellationTokenView


class TemporaryDirectoryUnavailable(OSError):
    """Private directory creation failed; no usable directory was published."""


def make_bounded_temp_dir(
    parent: str | os.PathLike[str],
    *,
    prefix: str,
    cancellation: CancellationTokenView | None = None,
    attempts: int = 16,
) -> str:
    """Create one random private directory, checking cancellation before retries."""
    if attempts < 1:
        raise ValueError("attempts must be positive")
    for _ in range(attempts):
        if cancellation is not None:
            cancellation.check()
        candidate = Path(parent) / f"{prefix}{secrets.token_hex(5)}"
        try:
            os.mkdir(candidate, 0o700)
        except FileExistsError:
            continue
        except PermissionError as exc:
            # Windows may report EACCES for an occupied name. Only that
            # precise case merits trying another random name.
            if os.path.lexists(candidate):
                continue
            raise TemporaryDirectoryUnavailable(
                "Cannot create a temporary work directory. Check folder permissions or security software."
            ) from exc
        except OSError as exc:
            raise TemporaryDirectoryUnavailable(
                "Cannot create a temporary work directory. Check folder permissions or available space."
            ) from exc
        return str(candidate)
    raise TemporaryDirectoryUnavailable("Cannot allocate a unique temporary work directory.")


@contextmanager
def bounded_temporary_directory(
    *,
    prefix: str,
    dir: str | os.PathLike[str],
    cancellation: CancellationTokenView | None = None,
) -> Iterator[str]:
    """Own and clean one temporary directory without tempfile's retry loop."""
    path = make_bounded_temp_dir(dir, prefix=prefix, cancellation=cancellation)
    try:
        yield path
    finally:
        shutil.rmtree(path)
