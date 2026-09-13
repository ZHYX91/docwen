"""Identity-bound operations on a Windows process, never on a recycled PID."""

from __future__ import annotations

import ctypes
import sys
from ctypes import wintypes
from dataclasses import dataclass
from typing import Any


def _kernel32() -> Any:
    if sys.platform != "win32":
        raise OSError("Windows process operations require Windows")
    library = ctypes.WinDLL("kernel32", use_last_error=True)
    library.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    library.OpenProcess.restype = wintypes.HANDLE
    library.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    library.GetProcessTimes.restype = wintypes.BOOL
    library.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    library.WaitForSingleObject.restype = wintypes.DWORD
    library.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    library.TerminateProcess.restype = wintypes.BOOL
    library.CloseHandle.argtypes = [wintypes.HANDLE]
    library.CloseHandle.restype = wintypes.BOOL
    return library


def _creation_time(library: Any, handle: Any) -> int | None:
    times = [wintypes.FILETIME() for _ in range(4)]
    if not library.GetProcessTimes(handle, *(ctypes.byref(value) for value in times)):
        return None
    return (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime


@dataclass(frozen=True)
class WindowsProcessIdentity:
    pid: int
    created_at: int

    @classmethod
    def capture(cls, pid: int, *, started_after_ns: int) -> WindowsProcessIdentity | None:
        """Claim only an explicitly identified process created during this attempt."""
        if sys.platform != "win32" or pid <= 0:
            return None
        library = _kernel32()
        handle = library.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED_INFORMATION
        if not handle:
            return None
        try:
            created_at = _creation_time(library, handle)
            boundary = started_after_ns // 100 + 116444736000000000
            if created_at is None or created_at < boundary:
                return None
            return cls(pid, created_at)
        finally:
            library.CloseHandle(handle)

    def terminate_if_running(self, *, grace_ms: int = 0) -> None:
        """Recheck identity and terminate through the same handle after a grace period."""
        if sys.platform != "win32":
            return
        library = _kernel32()
        handle = library.OpenProcess(0x1000 | 0x100000 | 0x1, False, self.pid)
        if not handle:
            return
        try:
            if _creation_time(library, handle) != self.created_at:
                return
            if library.WaitForSingleObject(handle, grace_ms) == 258:  # WAIT_TIMEOUT
                library.TerminateProcess(handle, 1)
        finally:
            library.CloseHandle(handle)
