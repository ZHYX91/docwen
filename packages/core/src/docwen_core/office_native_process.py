"""Bounded Windows Job / macOS process-group owners for external Office.

The parent supervisor in ``office_process`` accepts a result only after this
owner has exited. A launcher exit, a termination request, or closing a Job
handle alone never grants permission to delete the conversion workspace.
"""

from __future__ import annotations

import contextlib
import os
import signal
import sys
import tempfile
import time
from collections.abc import Callable
from typing import Any, BinaryIO, Protocol

from docwen_core.office_process import (
    _POLL_S,
    _REAP_TIMEOUT_S,
    _TERM_GRACE_S,
    OfficeProcessResult,
    OwnedProcessLaunch,
)


class NativeProcessFamily(Protocol):
    """One launch boundary with an independently observable exit receipt."""

    pid: int

    def start(
        self,
        command: list[str],
        *,
        stdin: int,
        stdout: int,
        cwd: str | None,
        environment: dict[str, str] | None,
    ) -> None: ...

    def exited(self) -> bool: ...

    def terminate(self, *, force: bool) -> None: ...

    def reap(self) -> int: ...

    def close(self) -> None: ...


def _native_family() -> NativeProcessFamily:
    if sys.platform == "win32":
        from docwen_core.windows_process import WindowsJobProcess

        return WindowsJobProcess()
    if sys.platform == "darwin":
        from docwen_core.macos_process import MacOSProcessGroup

        return MacOSProcessGroup()
    raise OSError("No native Office ownership boundary is available")


def _drain_family(
    family: NativeProcessFamily,
    *,
    is_stopped: Callable[[], bool],
    deadline: float,
    log: BinaryIO,
    log_limit: int,
    initial_error: str = "",
) -> tuple[bool, int | None, str]:
    """Use one finite closeout budget, including repeated observation failures."""
    reason = initial_error
    stopping_at = time.monotonic() if reason else None
    observation_failed = False
    while True:
        now = time.monotonic()
        try:
            if stopping_at is None:
                if is_stopped():
                    reason, stopping_at = "Office conversion cancelled", now
                elif now >= deadline:
                    reason, stopping_at = "Office conversion timed out", now
                elif log_limit and os.fstat(log.fileno()).st_size > log_limit:
                    reason, stopping_at = "Owned process diagnostic output exceeded its limit", now
            if family.exited():
                status = family.reap()
                # An ownership/identity failure is sticky. A later empty
                # snapshot cannot prove where a previously escaped child went.
                return not observation_failed, status, reason
            if stopping_at is not None:
                if now >= stopping_at + _TERM_GRACE_S + _REAP_TIMEOUT_S:
                    return False, None, reason + "; owned process exit could not be confirmed"
                family.terminate(force=now >= stopping_at + _TERM_GRACE_S)
        except BaseException as error:
            # Do not abandon a successfully created process on startup, query,
            # cancellation callback, or signal failures. Keep attempting bounded
            # cleanup, but never turn an uncertain ownership observation green.
            observation_failed = True
            if stopping_at is None:
                stopping_at = now
            reason = reason or f"Office process ownership failed: {type(error).__name__}"
            if now >= stopping_at + _TERM_GRACE_S + _REAP_TIMEOUT_S:
                return False, None, reason + "; owned process exit could not be confirmed"
            with contextlib.suppress(BaseException):
                family.terminate(force=now >= stopping_at + _TERM_GRACE_S)
        time.sleep(_POLL_S)


def native_owner_worker(
    connection: Any, stop: Any, command: list[str], timeout_s: float, launch: OwnedProcessLaunch | None = None
) -> None:
    """Create the boundary before business code, and report its confirmed exit."""
    family: NativeProcessFamily | None = None
    confirmed = False
    status: int | None = None
    reason = ""
    resources = contextlib.ExitStack()
    try:
        signal.signal(signal.SIGINT, signal.SIG_IGN)
        if stop.is_set():
            confirmed, reason = True, "cancelled before Office launch"
            return
        # Office never consumes stdout/stderr. No inherited PIPE can fill or
        # keep communicate() alive after its launcher exits. Other consumers
        # can opt into an owned, size-checked diagnostic file.
        log = resources.enter_context(open(launch.log_path if launch and launch.log_path else os.devnull, "wb"))  # noqa: SIM115
        stdin = resources.enter_context(tempfile.TemporaryFile("w+b"))  # noqa: SIM115
        if launch is not None and launch.stdin_bytes is not None:
            if len(launch.stdin_bytes) > 4096:
                raise ValueError("Owned process launch input exceeds its finite handshake budget")
            stdin.write(launch.stdin_bytes)
            stdin.seek(0)
        family = _native_family()
        try:
            family.start(
                command,
                stdin=stdin.fileno(),
                stdout=log.fileno(),
                cwd=launch.cwd if launch else None,
                environment=dict(launch.environment) if launch and launch.environment is not None else None,
            )
        except BaseException as error:
            reason = f"Office process launch failed: {type(error).__name__}"
        deadline = time.monotonic() + max(timeout_s, 0.0)
        with contextlib.suppress(OSError, EOFError):
            connection.send(("started", family.pid, deadline))
        confirmed, status, reason = _drain_family(
            family,
            is_stopped=stop.is_set,
            deadline=deadline,
            log=log,
            log_limit=launch.log_limit if launch else 0,
            initial_error=reason,
        )
    except BaseException as error:
        reason = f"Office native owner failed: {type(error).__name__}"
        confirmed = family is None
    finally:
        try:
            try:
                if family is not None:
                    family.close()
            finally:
                resources.close()
        except BaseException as error:
            confirmed = False
            reason = f"Office native owner closeout failed: {type(error).__name__}"
        with contextlib.suppress(OSError, EOFError):
            connection.send(
                OfficeProcessResult(
                    confirmed and not reason and status == 0,
                    confirmed,
                    reason,
                    os.getpid(),
                    family.pid if family else 0,
                    (family.pid,) if family and family.pid and status is not None else (),
                    launcher_exit_code=status,
                )
            )
        connection.close()
