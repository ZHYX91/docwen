"""Identity-pinned ownership of a controlled macOS Office process group.

Darwin has no Linux subreaper or kqueue NOTE_TRACK support. A dedicated owner
anchors a private session; its Office launcher starts a separate process group.
The launcher stays unreaped until all group members have stopped, pinning the
group ID while it is signalled. Observed members that escape the group invalidate
the proof. Arbitrary, unobserved daemonization is not a claimed containment API.
"""

from __future__ import annotations

import contextlib
import ctypes
import errno
import os
import signal
import subprocess
import sys
import time
from dataclasses import dataclass
from typing import Any

_MAX_GROUP_MEMBERS = 32768
_SCAN_TIMEOUT_S = 0.2
_SZOMB = 5


def _os_feature(name: str) -> Any:
    feature = getattr(os, name, None)
    if feature is None:
        raise OSError(f"macOS Office ownership requires {name}")
    return feature


class _SigValue(ctypes.Union):
    _fields_ = (("integer", ctypes.c_int32), ("pointer", ctypes.c_void_p))


class _SigInfo(ctypes.Structure):
    # Public 64-bit Darwin siginfo_t ABI (sys/signal.h), not Linux's layout.
    _fields_ = [
        ("signo", ctypes.c_int32),
        ("error", ctypes.c_int32),
        ("code", ctypes.c_int32),
        ("pid", ctypes.c_int32),
        ("uid", ctypes.c_uint32),
        ("status", ctypes.c_int32),
        ("address", ctypes.c_void_p),
        ("value", _SigValue),
        ("band", ctypes.c_int64),
        ("reserved", ctypes.c_uint64 * 7),
    ]


class MacChildWaiter:
    """Observe a wait-child without reaping it, including on Python 3.12.

    CPython exposes os.waitid on macOS only from 3.13. The public Darwin
    libSystem call is available on every supported macOS version. Always use
    its native ABI and flags, independent of the interpreter's optional API.
    """

    api: Any

    def __init__(self) -> None:
        if sys.platform != "darwin" or ctypes.sizeof(ctypes.c_void_p) != 8:
            raise OSError("Darwin wait-child observation requires 64-bit macOS")
        self.api = ctypes.CDLL("/usr/lib/libSystem.B.dylib", use_errno=True)
        self.api.waitid.argtypes = [ctypes.c_int, ctypes.c_uint32, ctypes.POINTER(_SigInfo), ctypes.c_int]
        self.api.waitid.restype = ctypes.c_int

    def exited(self, pid: int) -> bool:
        if pid <= 0:
            raise OSError("Cannot observe an invalid Office wait-child identity")
        deadline = time.monotonic() + _SCAN_TIMEOUT_S
        while True:
            # Darwin may leave siginfo untouched under WNOHANG; start cleared.
            info = _SigInfo()
            ctypes.set_errno(0)
            # sys/wait.h: P_PID=1, WEXITED=4, WNOHANG=1, WNOWAIT=0x20.
            result = self.api.waitid(1, pid, ctypes.byref(info), 0x25)
            if result == 0:
                break
            error = ctypes.get_errno()
            if error != errno.EINTR or time.monotonic() >= deadline:
                raise OSError(error, "Cannot establish the Darwin Office wait-child identity")
        if info.pid == 0:
            return False
        # sys/signal.h: SIGCHLD=20; CLD_EXITED/KILLED/DUMPED are 1/2/3.
        if info.pid != pid or info.signo != 20 or info.code not in (1, 2, 3):
            raise OSError("Darwin returned an invalid Office wait-child exit receipt")
        return True


class _ProcBSDInfo(ctypes.Structure):
    # Public Darwin proc_bsdinfo ABI (sys/proc_info.h), including start time.
    _fields_ = [
        (name, ctypes.c_uint32)
        for name in (
            "flags",
            "status",
            "exit_status",
            "pid",
            "ppid",
            "uid",
            "gid",
            "ruid",
            "rgid",
            "svuid",
            "svgid",
            "reserved",
        )
    ] + [
        ("comm", ctypes.c_char * 16),
        ("name", ctypes.c_char * 32),
        ("nfiles", ctypes.c_uint32),
        ("pgid", ctypes.c_uint32),
        ("jobc", ctypes.c_uint32),
        ("tty_dev", ctypes.c_uint32),
        ("tty_pgid", ctypes.c_uint32),
        ("nice", ctypes.c_int32),
        ("start_seconds", ctypes.c_uint64),
        ("start_microseconds", ctypes.c_uint64),
    ]


@dataclass(frozen=True, slots=True)
class MacProcessIdentity:
    pid: int
    group: int
    status: int
    started: tuple[int, int]


class MacProcessInventory:
    api: Any

    def __init__(self) -> None:
        if sys.platform != "darwin":
            raise OSError("Darwin process inventory requires macOS")
        self.api: Any = ctypes.CDLL("/usr/lib/libproc.dylib", use_errno=True)
        self.api.proc_listpids.argtypes = [ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p, ctypes.c_int]
        self.api.proc_listpids.restype = ctypes.c_int
        self.api.proc_pidinfo.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_uint64, ctypes.c_void_p, ctypes.c_int]
        self.api.proc_pidinfo.restype = ctypes.c_int
        own = self.identity(os.getpid())
        if own is None or own.group != _os_feature("getpgrp")():
            raise OSError("Cannot establish Darwin owner process identity")

    def identity(self, pid: int) -> MacProcessIdentity | None:
        info = _ProcBSDInfo()
        ctypes.set_errno(0)
        # PROC_PIDTBSDINFO requires a nonzero argument to include zombies.
        # The deliberately unreaped launcher must remain identity-observable.
        size = self.api.proc_pidinfo(pid, 3, 1, ctypes.byref(info), ctypes.sizeof(info))
        if size == 0 and ctypes.get_errno() == errno.ESRCH:
            return None
        if size != ctypes.sizeof(info) or info.pid != pid:
            raise OSError(ctypes.get_errno(), "Cannot establish a Darwin process identity")
        return MacProcessIdentity(pid, info.pgid, info.status, (info.start_seconds, info.start_microseconds))

    def group(self, group_id: int) -> tuple[MacProcessIdentity, ...]:
        deadline = time.monotonic() + _SCAN_TIMEOUT_S
        buffer = (ctypes.c_int * (_MAX_GROUP_MEMBERS + 1))()
        ctypes.set_errno(0)
        size = self.api.proc_listpids(2, group_id, buffer, ctypes.sizeof(buffer))  # PROC_PGRP_ONLY
        if size < 0 or (size == 0 and ctypes.get_errno()):
            raise OSError(ctypes.get_errno(), "Cannot enumerate the owned Darwin process group")
        if size >= ctypes.sizeof(buffer) or size % ctypes.sizeof(ctypes.c_int):
            raise OSError("Darwin process group exceeded its finite inventory budget")
        members: list[MacProcessIdentity] = []
        for pid in buffer[: size // ctypes.sizeof(ctypes.c_int)]:
            if time.monotonic() >= deadline:
                raise OSError("Darwin process inventory exceeded its finite scan deadline")
            if pid <= 0:
                continue
            identity = self.identity(pid)
            if identity is not None:
                if identity.group != group_id:
                    raise OSError("An observed Office process left its owned group")
                members.append(identity)
        return tuple(members)


class MacOSProcessGroup:
    def __init__(self) -> None:
        self.pid = 0
        self.process: subprocess.Popen[bytes] | None = None
        self.known: dict[int, MacProcessIdentity] = {}
        self.reaped = False
        for name in ("setsid", "killpg", "WNOHANG"):
            _os_feature(name)
        kill_signal = getattr(signal, "SIGKILL", None)
        if kill_signal is None:
            raise OSError("macOS Office ownership requires a force signal")
        self.kill_signal = kill_signal
        self.child_waiter = MacChildWaiter()
        # This runs only in the spawned owner, never in the GUI/CLI process.
        _os_feature("setsid")()
        self.inventory = MacProcessInventory()

    def start(
        self,
        command: list[str],
        *,
        stdin: int,
        stdout: int,
        cwd: str | None,
        environment: dict[str, str] | None,
    ) -> None:
        self.process = subprocess.Popen(
            command,
            stdin=stdin,
            stdout=stdout,
            stderr=subprocess.STDOUT,
            cwd=cwd,
            env=environment,
            process_group=0,
        )
        self.pid = self.process.pid
        identity = self.inventory.identity(self.pid)
        if identity is None or identity.group != self.pid:
            raise OSError("Cannot establish the private Office process group")
        self.known[self.pid] = identity

    def _peek_launcher(self) -> bool:
        # Never Popen.poll()/wait() here: reaping a fast-exiting launcher would
        # release its PID while a later killpg still targets that numeric ID.
        return self.child_waiter.exited(self.pid)

    def exited(self) -> bool:
        if self.process is None:
            return True
        deadline = time.monotonic() + _SCAN_TIMEOUT_S
        observation = self._peek_launcher()
        members = {member.pid: member for member in self.inventory.group(self.pid)}
        for pid, known in tuple(self.known.items()):
            if time.monotonic() >= deadline:
                raise OSError("Office process identity checks exceeded their finite deadline")
            current = self.inventory.identity(pid)
            if current is not None and current.started == known.started and current.group != self.pid:
                raise OSError("An owned Office process escaped its controlled group")
            if current is None or current.started != known.started:
                if pid == self.pid:
                    raise OSError("The unreaped Office launcher identity was lost")
                self.known.pop(pid)
            else:
                # A known live identity cannot disappear from the proof merely
                # because an earlier group snapshot omitted it during a race.
                members[pid] = current
        self.known.update(members)
        return observation and not any(member.status != _SZOMB for member in members.values())

    def terminate(self, *, force: bool) -> None:
        if self.process is None or self.reaped:
            return
        # A wait-child relation plus the unreaped launcher pins this group ID.
        # Failure to establish it grants no authority to signal a numeric PID.
        self._peek_launcher()
        # Only a subsequent complete inventory establishes cleanup.
        with contextlib.suppress(ProcessLookupError):
            _os_feature("killpg")(self.pid, self.kill_signal if force else signal.SIGTERM)

    def reap(self) -> int:
        if self.process is None:
            return -1
        if not self.exited():
            raise OSError("The controlled Office process group is still active")
        pid, status = os.waitpid(self.pid, _os_feature("WNOHANG"))
        if pid != self.pid:
            raise OSError("Cannot reap the owned Office launcher")
        self.reaped = True
        self.process.returncode = os.waitstatus_to_exitcode(status)
        return self.process.returncode

    def close(self) -> None:
        # Do not wait without a deadline or delete retained state. If closeout
        # was unconfirmed, the bridge returns a failure and keeps all inputs.
        pass
