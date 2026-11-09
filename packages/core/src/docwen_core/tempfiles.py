"""Shared bounded temporary-file and directory ownership for all DocWen layers."""

from __future__ import annotations

import os
import secrets
import shutil
import tempfile
from collections.abc import Callable
from pathlib import Path
from types import TracebackType


class TemporaryDirectoryError(OSError):
    """A private temporary directory could not be created."""


def make_temp_dir(
    suffix: str = "",
    prefix: str = "tmp",
    dir: str | os.PathLike[str] | None = None,
    *,
    cancellation_check: Callable[[], None] | None = None,
    attempts: int = 16,
) -> str:
    """Create a private directory; retry name collisions only, never persistent denial."""
    if attempts < 1:
        raise ValueError("attempts must be positive")
    parent = Path(dir) if dir is not None else Path(tempfile.gettempdir())
    for _ in range(attempts):
        if cancellation_check is not None:
            cancellation_check()
        candidate = parent / f"{prefix}{secrets.token_hex(5)}{suffix}"
        try:
            os.mkdir(candidate, 0o700)
        except FileExistsError:
            continue
        except PermissionError as exc:
            # On Windows, EACCES can also mean the randomly chosen name is
            # already occupied. Retry only when that exact entry exists.
            if os.path.lexists(candidate):
                continue
            raise TemporaryDirectoryError(
                "Cannot create a temporary work directory. Check folder permissions or security software."
            ) from exc
        except OSError as exc:
            raise TemporaryDirectoryError(
                "Cannot create a temporary work directory. Check folder permissions or available space."
            ) from exc
        return str(candidate)
    raise TemporaryDirectoryError("Cannot allocate a unique temporary work directory.")


class TemporaryDirectory:
    """Temporary directory owner with bounded creation and honest cleanup errors.

    If cleanup fails while another exception is already propagating, preserve
    that primary exception and attach cleanup information instead of masking it.
    """

    def __init__(
        self,
        suffix: str = "",
        prefix: str = "tmp",
        dir: str | os.PathLike[str] | None = None,
        *,
        ignore_cleanup_errors: bool = False,
        cancellation_check: Callable[[], None] | None = None,
    ) -> None:
        self.name = make_temp_dir(
            suffix=suffix,
            prefix=prefix,
            dir=dir,
            cancellation_check=cancellation_check,
        )
        self._ignore_cleanup_errors = ignore_cleanup_errors
        self._cleaned = False

    def cleanup(self) -> None:
        """Remove owned files, retaining ownership when cleanup needs a retry."""
        if self._cleaned:
            return
        try:
            shutil.rmtree(self.name)
        except OSError as exc:
            if isinstance(exc, FileNotFoundError):
                # A missing child is still a failed cleanup when the owned
                # directory remains. Only an absent root completes ownership.
                try:
                    os.lstat(self.name)
                except FileNotFoundError:
                    self._cleaned = True
                    return
                except OSError:
                    pass
            if not self._ignore_cleanup_errors:
                raise
            return
        self._cleaned = True

    def __enter__(self) -> str:
        return self.name

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        try:
            self.cleanup()
        except OSError as cleanup_error:
            if exc_value is None:
                raise
            exc_value.add_note(f"Temporary directory cleanup failed: {type(cleanup_error).__name__}")
