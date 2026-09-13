"""Read-only process observations for engineering lease ownership."""

from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass


@dataclass(frozen=True)
class ProcessObservation:
    alive: bool
    identity: str | None = None


def observe(pid: object) -> ProcessObservation:
    if not isinstance(pid, int) or pid <= 0:
        return ProcessObservation(False)
    if os.name == "nt":
        return _windows_observe(pid)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return ProcessObservation(False)
    except PermissionError:
        return ProcessObservation(True)
    except OSError:
        return ProcessObservation(False)
    return ProcessObservation(True)


def owner_alive(pid: object, identity: object) -> bool:
    observation = observe(pid)
    if not observation.alive:
        return False
    if not isinstance(identity, str) or re.fullmatch(r"windows-filetime:[0-9a-f]{16}", identity) is None:
        return True
    return observation.identity is None or observation.identity == identity


def _windows_observe(pid: int) -> ProcessObservation:
    # os.kill(pid, 0) sends CTRL_C_EVENT on Windows. Never use it here.
    if sys.platform != "win32":
        raise OSError("Windows process handles are unavailable on this platform")
    import ctypes
    from ctypes import wintypes

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
    kernel.WaitForSingleObject.restype = wintypes.DWORD
    kernel.GetProcessTimes.argtypes = (wintypes.HANDLE,) + (ctypes.POINTER(wintypes.FILETIME),) * 4
    kernel.GetProcessTimes.restype = wintypes.BOOL
    kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel.CloseHandle.restype = wintypes.BOOL
    handle = kernel.OpenProcess(0x00101000, False, pid)  # SYNCHRONIZE | QUERY_LIMITED_INFORMATION
    if not handle:
        return ProcessObservation(ctypes.get_last_error() != 87)
    try:
        wait = kernel.WaitForSingleObject(handle, 0)
        if wait == 0:
            return ProcessObservation(False)
        if wait != 258:
            return ProcessObservation(True)
        creation, end, system, user = (wintypes.FILETIME() for _ in range(4))
        if not kernel.GetProcessTimes(
            handle, ctypes.byref(creation), ctypes.byref(end), ctypes.byref(system), ctypes.byref(user)
        ):
            return ProcessObservation(True)
        ticks = (creation.dwHighDateTime << 32) | creation.dwLowDateTime
        return ProcessObservation(True, f"windows-filetime:{ticks:016x}")
    finally:
        kernel.CloseHandle(handle)
