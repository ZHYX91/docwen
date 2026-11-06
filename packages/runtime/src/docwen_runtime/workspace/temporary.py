"""Runtime-local import bridge to Core's shared temporary directory policy.

All creation and cleanup behavior lives in docwen_core.tempfiles. This module
keeps Runtime call sites readable while the migration is completed.
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from os import PathLike
from typing import TYPE_CHECKING

from docwen_core.tempfiles import TemporaryDirectory, make_temp_dir

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import CancellationTokenView


def make_bounded_temp_dir(
    parent: str | PathLike[str],
    *,
    prefix: str,
    cancellation: CancellationTokenView | None = None,
    attempts: int = 16,
) -> str:
    return make_temp_dir(
        prefix=prefix,
        dir=parent,
        cancellation_check=cancellation.check if cancellation is not None else None,
        attempts=attempts,
    )


@contextmanager
def bounded_temporary_directory(
    *,
    prefix: str,
    dir: str | PathLike[str],
    cancellation: CancellationTokenView | None = None,
) -> Iterator[str]:
    with TemporaryDirectory(
        prefix=prefix,
        dir=dir,
        cancellation_check=cancellation.check if cancellation is not None else None,
    ) as path:
        yield path
