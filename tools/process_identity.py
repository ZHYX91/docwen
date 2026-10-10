"""Read-only process observations for engineering lease ownership."""

from __future__ import annotations

import json
import os
import re
import subprocess
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
    if observation.identity and observation.identity.startswith("windows-cim-filetime:"):
        # CIM timestamps have microsecond precision; never interpret rounding
        # within that interval as PID reuse.
        ticks = int(observation.identity.split(":", 1)[1], 16)
        return abs(int(identity.split(":", 1)[1], 16) - ticks) < 10
    return observation.identity is None or observation.identity == identity


def _cim_process_identity(pid: int) -> str | None:
    """Read process birth without opening a protected process handle."""
    if not isinstance(pid, int) or isinstance(pid, bool) or pid <= 0:
        return None
    script = (
        "$ErrorActionPreference='Stop'; "
        f"$p=Get-CimInstance Win32_Process -Filter 'ProcessId={pid}'; "
        "if ($null -ne $p.CreationDate) { "
        "@{pid=$p.ProcessId;ticks=$p.CreationDate.ToUniversalTime().ToFileTimeUtc().ToString()} "
        "| ConvertTo-Json -Compress }"
    )
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
        if result.returncode:
            return None
        record = json.loads(result.stdout)
        if not isinstance(record, dict) or record.get("pid") != pid:
            return None
        value = record.get("ticks")
        if not isinstance(value, str) or not re.fullmatch(r"[0-9]{1,19}", value):
            return None
        ticks = int(value)
        if not 0 < ticks < 2**63:
            return None
        return f"windows-cim-filetime:{ticks:016x}"
    except (OSError, subprocess.TimeoutExpired, ValueError):
        return None


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
        error = ctypes.get_last_error()
        return ProcessObservation(error != 87, _cim_process_identity(pid) if error == 5 else None)
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
