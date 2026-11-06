"""Bounded retries for replacing a fully prepared private file on Windows.

Callers must first authorize the destination; this operation overwrites and
must never be used in a no-clobber publication path.
"""

from __future__ import annotations

import os
import sys
import time
from collections.abc import Callable
from pathlib import Path

_TRANSIENT_WINERRORS = frozenset({5, 32})


def _is_windows() -> bool:
    return sys.platform == "win32"


def replace_authorized_file(
    prepared: str | os.PathLike[str],
    destination: str | os.PathLike[str],
    *,
    before_attempt: Callable[[], None] | None = None,
    attempts: int = 5,
) -> None:
    """Atomic replacement: retry only transient Windows access/sharing errors."""
    if attempts < 1:
        raise ValueError("attempts must be positive")
    source = Path(prepared)
    target = Path(destination)
    for attempt in range(attempts):
        if before_attempt is not None:
            before_attempt()
        try:
            os.replace(source, target)
            return
        except OSError as exc:
            if (
                not _is_windows()
                or getattr(exc, "winerror", None) not in _TRANSIENT_WINERRORS
                or attempt == attempts - 1
            ):
                raise
            time.sleep(0.02 * (attempt + 1))
