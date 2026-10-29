"""Linux-only ownership of one external Office process family.

The spawned owner is a local subreaper. It never changes the application's
process-wide child ownership and never treats a launcher's exit as quiescence.
"""

from __future__ import annotations

import contextlib
import ctypes
import multiprocessing
import os
import signal
import subprocess
import sys
import time
from collections.abc import Callable
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

_POLL_S = 0.02
_STARTUP_TIMEOUT_S = 5.0
_TERM_GRACE_S = 0.4
_REAP_TIMEOUT_S = 3.0
_OWNER_JOIN_S = 1.0
_PROC_ROOT = Path("/proc")
_TASK_ROOT = _PROC_ROOT / "self/task"
_DISCOVERY_TIMEOUT_S = 0.2
_DISCOVERY_MAX_ENTRIES = 32768
_PROC_TEXT_MAX_CHARS = 65536


@dataclass(frozen=True, slots=True)
class OfficeProcessResult:
    succeeded: bool
    cleanup_confirmed: bool
    reason: str = ""
    owner_pid: int = 0
    launcher_pid: int = 0
    reaped_pids: tuple[int, ...] = ()
    owner_exit_code: int | None = None
    launcher_exit_code: int | None = None


class OfficeProcessCleanupUnconfirmed(RuntimeError):
    """An interrupted wait could not prove that owned processes were reaped."""


def _proc_text(path: Path) -> str:
    with path.open(encoding="utf-8", errors="replace") as stream:
        text = stream.read(_PROC_TEXT_MAX_CHARS + 1)
    if len(text) > _PROC_TEXT_MAX_CHARS:
        raise OSError("Office process discovery exceeded its finite text budget")
    return text


def _namespace_ids(status: str) -> tuple[int, ...]:
    rows = [line.split()[1:] for line in status.splitlines() if line.startswith("NSpid:")]
    if len(rows) != 1 or not rows[0]:
        raise OSError("Cannot establish procfs PID namespace identity")
    try:
        identities = tuple(int(value) for value in rows[0])
    except ValueError as error:
        raise OSError("Invalid procfs PID namespace identity") from error
    if any(pid <= 0 for pid in identities):
        raise OSError("Invalid procfs PID namespace identity")
    return identities


def _namespace_pids(status_path: Path) -> tuple[int, ...]:
    return _namespace_ids(_proc_text(status_path))


def _discovery_budget(index: int, deadline: float) -> None:
    if index >= _DISCOVERY_MAX_ENTRIES or time.monotonic() >= deadline:
        raise OSError("Office direct-child discovery exceeded its finite scan budget")


def _task_child_proc_pids(owner_proc_pid: int, deadline: float) -> set[int] | None:
    """Prefer task children files, without treating unavailable files as empty."""
    children: set[int] = set()
    readable = 0
    unavailable = False
    try:
        with os.scandir(_TASK_ROOT) as entries:
            for index, entry in enumerate(entries):
                _discovery_budget(index, deadline)
                if not entry.name.isdecimal():
                    continue
                task = Path(entry.path)
                try:
                    values = _proc_text(task / "children").split()
                except (FileNotFoundError, PermissionError):
                    # A disappearing secondary thread may be skipped. Our own
                    # main thread cannot have exited while executing this code;
                    # its missing children file is an unavailable interface.
                    if int(entry.name) == owner_proc_pid or task.exists():
                        unavailable = True
                    continue
                readable += 1
                try:
                    pids = {int(value) for value in values}
                except ValueError as error:
                    raise OSError("Invalid Office task child identity") from error
                if any(pid <= 0 for pid in pids):
                    raise OSError("Invalid Office task child identity")
                children.update(pids)
    except (FileNotFoundError, PermissionError):
        return None
    return None if unavailable or not readable else children


def _mapped_child(proc_pid: int, identities: tuple[int, ...], local_depth: int) -> int:
    if identities[0] != proc_pid or len(identities) <= local_depth:
        raise OSError("Cannot map an Office child's PID namespace")
    return identities[local_depth]


def _status_direct_children(owner_proc_pid: int, local_depth: int, deadline: float) -> set[int]:
    """Fallback discovery only: PPid is in the mounted procfs PID namespace."""
    children: set[int] = set()
    with os.scandir(_PROC_ROOT) as entries:
        for index, entry in enumerate(entries):
            _discovery_budget(index, deadline)
            if not entry.name.isdecimal():
                continue
            try:
                status = _proc_text(Path(entry.path) / "status")
            except (FileNotFoundError, ProcessLookupError, PermissionError):
                # A gone/inaccessible process supplies no candidate or signal
                # authority. Only kernel ECHILD can prove final quiescence.
                continue
            rows = [line.split()[1:] for line in status.splitlines() if line.startswith("PPid:")]
            if len(rows) != 1 or len(rows[0]) != 1:
                continue
            try:
                parent = int(rows[0][0])
            except ValueError:
                continue
            if parent == owner_proc_pid:
                children.add(_mapped_child(int(entry.name), _namespace_ids(status), local_depth))
    return children


def _direct_children() -> set[int]:
    # Procfs may be mounted in an ancestor PID namespace. Its children/status
    # numbers cannot be sent directly to waitid or pidfd_open in our namespace.
    own_namespaces = _namespace_pids(_PROC_ROOT / "self/status")
    if own_namespaces[-1] != os.getpid():
        raise OSError("Cannot map the Office owner's PID namespace")
    local_depth = len(own_namespaces) - 1
    deadline = time.monotonic() + _DISCOVERY_TIMEOUT_S
    proc_pids = _task_child_proc_pids(own_namespaces[0], deadline)
    if proc_pids is None:
        return _status_direct_children(own_namespaces[0], local_depth, deadline)
    children: set[int] = set()
    for index, proc_pid in enumerate(proc_pids):
        _discovery_budget(index, deadline)
        try:
            identities = _namespace_pids(_PROC_ROOT / str(proc_pid) / "status")
        except (FileNotFoundError, ProcessLookupError):
            continue
        children.add(_mapped_child(proc_pid, identities, local_depth))
    # Discovery still grants no signal authority; _live_direct_child proves
    # the mapped local identity against the kernel before opening each pidfd.
    return children


def _live_direct_child(pid: int) -> bool:
    """Keep procfs discovery subordinate to the kernel wait-child relation."""
    waitid = getattr(os, "waitid", None)
    p_pid = getattr(os, "P_PID", None)
    wexited = getattr(os, "WEXITED", None)
    wnowait = getattr(os, "WNOWAIT", None)
    wnohang = getattr(os, "WNOHANG", None)
    if waitid is None or p_pid is None or wexited is None or wnowait is None or wnohang is None:
        raise OSError("Office ownership requires Linux wait-child support")
    try:
        observation = waitid(p_pid, pid, wexited | wnohang | wnowait)
    except ChildProcessError:
        return False
    # CPython returns None for a live child with WNOHANG. An exit observation
    # remains unreaped for the single waitpid authority in the owner loop.
    return observation is None or observation.si_pid == 0


def _enable_owner() -> None:
    if sys.platform != "linux":
        raise OSError("Linux Office ownership is unavailable")
    library = ctypes.CDLL(None, use_errno=True)
    prctl = library.prctl
    prctl.argtypes = [ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong]
    prctl.restype = ctypes.c_int
    if prctl(36, 1, 0, 0, 0) != 0:  # PR_SET_CHILD_SUBREAPER
        raise OSError(ctypes.get_errno(), "Cannot establish Office child ownership")
    open_pidfd = getattr(os, "pidfd_open", None)
    send_pidfd = getattr(signal, "pidfd_send_signal", None)
    if open_pidfd is None or send_pidfd is None:
        raise OSError("Office ownership requires Linux pidfd support")
    # Probe availability without spawning any business process. This is not a
    # process-exit test; only waitpid/ECHILD establishes the final proof.
    descriptor = open_pidfd(os.getpid(), 0)
    os.close(descriptor)
    _direct_children()
    signal.signal(signal.SIGINT, signal.SIG_IGN)


def _owner_worker(connection: Any, stop: Any, cmd: list[str], timeout_s: float) -> None:
    launcher: subprocess.Popen[bytes] | None = None
    descriptors: dict[int, int] = {}
    reaped: list[int] = []
    launcher_status: int | None = None
    reason = ""
    stopping_at: float | None = None
    deadline = float("inf")
    cleanup_confirmed = False
    try:
        _enable_owner()
        wnohang = getattr(os, "WNOHANG", None)
        if wnohang is None:
            raise OSError("Office ownership requires nonblocking Linux wait support")
        if stop.is_set():
            cleanup_confirmed = True
            reason = "cancelled before Office launch"
            return
        launcher = subprocess.Popen(
            cmd, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True
        )
        deadline = time.monotonic() + max(timeout_s, 0.0)
        with contextlib.suppress(OSError, EOFError):
            connection.send(("started", launcher.pid, deadline))
        while True:
            now = time.monotonic()
            if stopping_at is None and (stop.is_set() or now >= deadline):
                stopping_at = now
                reason = "Office conversion timed out" if now >= deadline else "Office conversion cancelled"
            # This is the only waiter for this process family. In particular,
            # Popen.poll/communicate cannot steal the launcher's exit status.
            while True:
                try:
                    child, status = os.waitpid(-1, wnohang)
                except ChildProcessError:
                    cleanup_confirmed = True
                    break
                if child == 0:
                    break
                reaped.append(child)
                if child == launcher.pid:
                    launcher_status = os.waitstatus_to_exitcode(status)
                    launcher.returncode = launcher_status
                descriptor = descriptors.pop(child, None)
                if descriptor is not None:
                    os.close(descriptor)
            if cleanup_confirmed:
                break
            now = time.monotonic()
            if stopping_at is not None:
                if now >= stopping_at + _TERM_GRACE_S + _REAP_TIMEOUT_S:
                    reason += "; owned process exit/reap could not be confirmed"
                    break
                force = now >= stopping_at + _TERM_GRACE_S
                try:
                    open_pidfd = getattr(os, "pidfd_open", None)
                    send_pidfd = getattr(signal, "pidfd_send_signal", None)
                    kill_signal = getattr(signal, "SIGKILL", None)
                    if open_pidfd is None or send_pidfd is None or kill_signal is None:
                        raise OSError("Office ownership requires Linux pidfd signaling")
                    for child in _direct_children():
                        if not _live_direct_child(child):
                            continue
                        descriptor = descriptors.get(child)
                        if descriptor is None:
                            # No other thread reaps children. Until this owner
                            # waits them, a discovered child's PID cannot be reused.
                            try:
                                descriptor = open_pidfd(child, 0)
                            except ProcessLookupError:
                                continue
                            descriptors[child] = descriptor
                        with contextlib.suppress(ProcessLookupError):
                            send_pidfd(descriptor, kill_signal if force else signal.SIGTERM)
                except OSError as error:
                    reason = f"Office process termination could not be confirmed: {type(error).__name__}"
            time.sleep(_POLL_S)
    except BaseException as error:
        reason = f"Office process owner failed: {type(error).__name__}"
        # Before Popen returned, no Office child exists. After it returned an
        # exception cannot authorize deletion; the parent retains its workspace.
        cleanup_confirmed = launcher is None
    finally:
        for descriptor in descriptors.values():
            with contextlib.suppress(OSError):
                os.close(descriptor)
        with contextlib.suppress(OSError, EOFError):
            connection.send(
                OfficeProcessResult(
                    cleanup_confirmed and stopping_at is None and launcher_status == 0,
                    cleanup_confirmed,
                    reason,
                    os.getpid(),
                    launcher.pid if launcher is not None else 0,
                    tuple(reaped),
                    launcher_exit_code=launcher_status,
                )
            )
        connection.close()


def _start_owner(cmd: list[str], timeout_s: float) -> tuple[Any, Any, Any]:
    context = multiprocessing.get_context("spawn")
    parent_connection, child_connection = context.Pipe(duplex=False)
    stop = context.Event()
    owner = context.Process(target=_owner_worker, args=(child_connection, stop, cmd, timeout_s), daemon=True)
    try:
        owner.start()
    except BaseException:
        parent_connection.close()
        child_connection.close()
        raise
    child_connection.close()
    return owner, parent_connection, stop


def run_owned_office(
    cmd: list[str],
    *,
    is_cancelled: Callable[[], bool],
    timeout_s: float,
) -> OfficeProcessResult:
    """Keep every interrupted/failed control-plane exit explicitly unconfirmed."""
    try:
        return _run_owned_office(cmd, is_cancelled=is_cancelled, timeout_s=timeout_s)
    except OfficeProcessCleanupUnconfirmed:
        raise
    except BaseException as error:
        # A second interrupt can arrive even while recording the first one or
        # joining/closing the owner. It must not escape as an ordinary exception
        # that the bridge's finally would mistake for cleanup permission.
        raise OfficeProcessCleanupUnconfirmed(
            f"Office process cleanup proof was interrupted: {type(error).__name__}"
        ) from error


def _run_owned_office(
    cmd: list[str],
    *,
    is_cancelled: Callable[[], bool],
    timeout_s: float,
) -> OfficeProcessResult:
    if is_cancelled() or timeout_s <= 0:
        return OfficeProcessResult(False, True, "cancelled before Office launch")
    try:
        owner, connection, stop = _start_owner(cmd, timeout_s)
    except Exception as error:
        return OfficeProcessResult(False, False, f"Cannot confirm Office owner startup: {type(error).__name__}")
    result: OfficeProcessResult | None = None
    business_deadline: float | None = None
    startup_deadline = time.monotonic() + _STARTUP_TIMEOUT_S
    stop_deadline: float | None = None
    interruption: BaseException | None = None
    while True:
        try:
            now = time.monotonic()
            if is_cancelled() or now >= (business_deadline if business_deadline is not None else startup_deadline):
                if stop_deadline is None:
                    stop_deadline = now + _TERM_GRACE_S + _REAP_TIMEOUT_S + _OWNER_JOIN_S
                stop.set()
            if connection.poll(_POLL_S):
                try:
                    message = connection.recv()
                except (EOFError, OSError):
                    break
                if isinstance(message, OfficeProcessResult):
                    result = message
                    break
                if isinstance(message, tuple) and len(message) == 3 and message[0] == "started":
                    business_deadline = float(message[2])
            if not owner.is_alive() and not connection.poll():
                break
        except BaseException as error:
            if interruption is None:
                interruption = error
            else:
                interruption.add_note(f"Office closeout also interrupted by {type(error).__name__}")
            # This budget is established only once. Repeated Ctrl+C or failing
            # pipe operations cannot reset it or turn absent proof into success.
            if stop_deadline is None:
                stop_deadline = time.monotonic() + _TERM_GRACE_S + _REAP_TIMEOUT_S + _OWNER_JOIN_S
            try:
                stop.set()
            except BaseException as stop_error:
                interruption.add_note(f"Office cancellation signal failed: {type(stop_error).__name__}")
        if stop_deadline is not None and time.monotonic() >= stop_deadline:
            break
    try:
        connection.close()
    except BaseException as error:
        interruption = interruption or error
    try:
        owner.join(timeout=_OWNER_JOIN_S)
    except BaseException as error:
        interruption = interruption or error
    owner_pid = owner.pid or 0
    owner_alive = owner.is_alive()
    owner_exit_code = owner.exitcode
    if result is None or owner_alive or owner_exit_code != 0 or result.owner_pid != owner_pid:
        outcome = OfficeProcessResult(
            False,
            False,
            "Office owner did not exit normally with a complete matching reap receipt",
            owner_pid,
            owner_exit_code=owner_exit_code,
        )
    else:
        outcome = replace(result, owner_exit_code=owner_exit_code)
    if not owner_alive:
        owner.close()
    if interruption is not None:
        raise OfficeProcessCleanupUnconfirmed(outcome.reason or "Office cleanup was interrupted") from interruption
    return outcome
