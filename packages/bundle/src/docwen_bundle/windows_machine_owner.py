"""Windows Machine process supervision inside the installed DocWenCLI.

The supervisor stays outside a private kill-on-close Job. Only the same CLI's
stdio worker is started, atomically assigned before its first instruction.
No executable is extracted, no process name is searched, and a parent PID is
used only to obtain a stable liveness handle, never as termination authority.
"""

from __future__ import annotations

import ctypes
import errno
import os
import subprocess
import sys
import time
from ctypes import wintypes as w
from typing import Any

OWNER_FAILURE = 125
TARGET_NOT_FOUND = 126
TARGET_FAILURE = 127
TARGET_CREATE_FAILED = 128
READY_RECORD = "DOCWEN_PROCESS_OWNER_READY:1\n"


class _Limits(ctypes.Structure):
    _fields_ = [
        ("process_time", ctypes.c_int64),
        ("job_time", ctypes.c_int64),
        ("flags", w.DWORD),
        ("min_working_set", ctypes.c_size_t),
        ("max_working_set", ctypes.c_size_t),
        ("active_limit", w.DWORD),
        ("affinity", ctypes.c_size_t),
        ("priority", w.DWORD),
        ("scheduling", w.DWORD),
    ]


class _ExtendedLimits(ctypes.Structure):
    _fields_ = [
        ("basic", _Limits),
        ("io", ctypes.c_uint64 * 6),
        ("process_memory", ctypes.c_size_t),
        ("job_memory", ctypes.c_size_t),
        ("peak_process_memory", ctypes.c_size_t),
        ("peak_job_memory", ctypes.c_size_t),
    ]


class _Accounting(ctypes.Structure):
    _fields_ = [
        ("times", ctypes.c_int64 * 4),
        ("page_faults", w.DWORD),
        ("total", w.DWORD),
        ("active", w.DWORD),
        ("terminated", w.DWORD),
    ]


class _Startup(ctypes.Structure):
    _fields_ = [
        ("cb", w.DWORD),
        ("reserved", w.LPWSTR),
        ("desktop", w.LPWSTR),
        ("title", w.LPWSTR),
        ("x", w.DWORD),
        ("y", w.DWORD),
        ("width", w.DWORD),
        ("height", w.DWORD),
        ("x_chars", w.DWORD),
        ("y_chars", w.DWORD),
        ("fill", w.DWORD),
        ("flags", w.DWORD),
        ("show", w.WORD),
        ("reserved_size", w.WORD),
        ("reserved_ptr", ctypes.c_void_p),
        ("stdin", w.HANDLE),
        ("stdout", w.HANDLE),
        ("stderr", w.HANDLE),
    ]


class _StartupEx(ctypes.Structure):
    _fields_ = [("startup", _Startup), ("attributes", ctypes.c_void_p)]


class _Process(ctypes.Structure):
    _fields_ = [("process", w.HANDLE), ("thread", w.HANDLE), ("pid", w.DWORD), ("tid", w.DWORD)]


def _kernel() -> Any:
    if sys.platform != "win32":
        raise OSError("Managed Windows stdio requires Windows")
    api = ctypes.WinDLL("kernel32", use_last_error=True)
    pointer = ctypes.c_void_p
    signatures = {
        "CreateJobObjectW": ([pointer, w.LPCWSTR], w.HANDLE),
        "SetInformationJobObject": ([w.HANDLE, ctypes.c_int, pointer, w.DWORD], w.BOOL),
        "QueryInformationJobObject": ([w.HANDLE, ctypes.c_int, pointer, w.DWORD, pointer], w.BOOL),
        "InitializeProcThreadAttributeList": ([pointer, w.DWORD, w.DWORD, pointer], w.BOOL),
        "UpdateProcThreadAttribute": (
            [pointer, w.DWORD, ctypes.c_size_t, pointer, ctypes.c_size_t, pointer, pointer],
            w.BOOL,
        ),
        "DeleteProcThreadAttributeList": ([pointer], None),
        "CreateProcessW": (
            [w.LPCWSTR, w.LPWSTR, pointer, pointer, w.BOOL, w.DWORD, pointer, w.LPCWSTR, pointer, pointer],
            w.BOOL,
        ),
        "GetCurrentProcess": ([], w.HANDLE),
        "GetStdHandle": ([w.DWORD], w.HANDLE),
        "SetStdHandle": ([w.DWORD, w.HANDLE], w.BOOL),
        "DuplicateHandle": ([w.HANDLE, w.HANDLE, w.HANDLE, pointer, w.DWORD, w.BOOL, w.DWORD], w.BOOL),
        "OpenProcess": ([w.DWORD, w.BOOL, w.DWORD], w.HANDLE),
        "GetProcessTimes": ([w.HANDLE, pointer, pointer, pointer, pointer], w.BOOL),
        "WaitForMultipleObjects": ([w.DWORD, pointer, w.BOOL, w.DWORD], w.DWORD),
        "WaitForSingleObject": ([w.HANDLE, w.DWORD], w.DWORD),
        "GetExitCodeProcess": ([w.HANDLE, pointer], w.BOOL),
        "ResumeThread": ([w.HANDLE], w.DWORD),
        "TerminateJobObject": ([w.HANDLE, w.UINT], w.BOOL),
        "CloseHandle": ([w.HANDLE], w.BOOL),
    }
    for name, (args, result) in signatures.items():
        function = getattr(api, name)
        function.argtypes = args
        function.restype = result
    return api


def _require(result: Any) -> None:
    if not result:
        raise _windows_error()


def _last_error() -> int:
    if sys.platform != "win32":
        raise OSError("Windows error state requires Windows")
    return ctypes.get_last_error()


def _windows_error(code: int | None = None) -> OSError:
    if sys.platform != "win32":
        return OSError("Windows process operations require Windows")
    return ctypes.WinError(_last_error() if code is None else code)


def _created_at(api: Any, handle: Any) -> int:
    times = [w.FILETIME() for _ in range(4)]
    _require(api.GetProcessTimes(handle, *(ctypes.byref(item) for item in times)))
    return (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime


def _finish_job(api: Any, job: Any) -> bool:
    if not api.TerminateJobObject(job, OWNER_FAILURE):
        return False
    deadline = time.monotonic() + 2
    while True:
        accounting = _Accounting()
        if not api.QueryInformationJobObject(job, 1, ctypes.byref(accounting), ctypes.sizeof(accounting), None):
            return False
        if accounting.active == 0:
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.01)


def run_job_worker(command: list[str], *, parent_pid: int) -> int:
    """Run a fixed worker command under Job ownership (also used by native tests)."""
    api = _kernel()
    job = parent = None
    child = _Process()
    attributes = None
    attributes_initialized = False
    duplicates: list[Any] = []
    exit_code = OWNER_FAILURE
    try:
        if not 0 < parent_pid < 2**32 or parent_pid == os.getpid():
            raise OSError(87, "Invalid parent identity")
        parent = api.OpenProcess(0x100000 | 0x1000, False, parent_pid)
        _require(parent)
        current = api.GetCurrentProcess()
        if _created_at(api, parent) > _created_at(api, current) or api.WaitForSingleObject(parent, 0) != 258:
            raise OSError(87, "Parent identity is no longer alive")
        job = api.CreateJobObjectW(None, None)
        _require(job)
        limits = _ExtendedLimits()
        limits.basic.flags = 0x2000  # KILL_ON_JOB_CLOSE; breakaway is never allowed.
        _require(api.SetInformationJobObject(job, 9, ctypes.byref(limits), ctypes.sizeof(limits)))
        for standard in (-10, -11, -12):
            duplicate = w.HANDLE()
            _require(
                api.DuplicateHandle(current, api.GetStdHandle(standard), current, ctypes.byref(duplicate), 0, True, 2)
            )
            duplicates.append(duplicate.value)
        handles = (w.HANDLE * 3)(*duplicates)
        jobs = (w.HANDLE * 1)(job)
        size = ctypes.c_size_t()
        api.InitializeProcThreadAttributeList(None, 2, 0, ctypes.byref(size))
        if not 0 < size.value <= 65536:
            raise OSError(87, "Invalid process attribute size")
        attributes = ctypes.create_string_buffer(size.value)
        _require(api.InitializeProcThreadAttributeList(attributes, 2, 0, ctypes.byref(size)))
        attributes_initialized = True
        _require(api.UpdateProcThreadAttribute(attributes, 0, 0x2000D, jobs, ctypes.sizeof(jobs), None, None))
        _require(api.UpdateProcThreadAttribute(attributes, 0, 0x20002, handles, ctypes.sizeof(handles), None, None))
        startup = _StartupEx()
        startup.startup.cb = ctypes.sizeof(startup)
        startup.startup.flags = 0x100
        startup.startup.stdin, startup.startup.stdout, startup.startup.stderr = duplicates
        startup.attributes = ctypes.cast(attributes, ctypes.c_void_p)
        deadline = time.monotonic() + 5
        while True:
            if api.WaitForSingleObject(parent, 0) != 258:
                raise OSError(87, "Parent exited before worker creation")
            line = ctypes.create_unicode_buffer(subprocess.list2cmdline(command))
            if api.CreateProcessW(
                command[0], line, None, None, True, 0x08080004, None, None, ctypes.byref(startup), ctypes.byref(child)
            ):
                break
            error = _last_error()
            if error not in (32, 33) or time.monotonic() >= deadline:
                raise _windows_error(error)
            time.sleep(0.1)
        for duplicate in duplicates:
            _require(api.CloseHandle(duplicate))
        duplicates.clear()
        # The worker must be the only reader, otherwise a write after worker
        # death can succeed into a pipe that the supervisor never drains.
        if sys.stdin is not None:
            input_fd = sys.stdin.fileno()
            sys.stdin.close()
            # Python's original standard streams use closefd=False. Closing
            # their wrapper alone leaves the supervisor's pipe reader alive.
            try:
                os.close(input_fd)
            except OSError as error:
                if error.errno != errno.EBADF:
                    raise
        _require(api.SetStdHandle(-10, None))
        sys.stderr.write(READY_RECORD)
        sys.stderr.flush()
        if api.ResumeThread(child.thread) == 0xFFFFFFFF:
            raise _windows_error()
        watched = (w.HANDLE * 2)(parent, child.process)
        signalled = api.WaitForMultipleObjects(2, watched, False, 0xFFFFFFFF)
        if signalled == 1:
            result = w.DWORD()
            _require(api.GetExitCodeProcess(child.process, ctypes.byref(result)))
            exit_code = TARGET_FAILURE if result.value in (125, 126, 128) else int(result.value)
        elif signalled != 0:
            raise _windows_error()
    except (OSError, ValueError) as error:
        code = getattr(error, "winerror", None) or getattr(error, "errno", None) or 31
        if not child.process:
            exit_code = TARGET_NOT_FOUND if code in (2, 3) else TARGET_CREATE_FAILED
            sys.stderr.write(f"DOCWEN_OWNER_CREATE_ERROR:{code}\n")
        else:
            sys.stderr.write(f"DOCWEN_OWNER_ERROR:{code}\n")
        sys.stderr.flush()
    finally:
        if job and not _finish_job(api, job):
            exit_code = OWNER_FAILURE
        if attributes_initialized:
            api.DeleteProcThreadAttributeList(attributes)
        for handle in [*duplicates, child.thread, child.process, job, parent]:
            if handle and not api.CloseHandle(handle):
                exit_code = OWNER_FAILURE
    return exit_code


def run_managed_stdio(parent_pid: int) -> int:
    command = [sys.executable]
    if not getattr(sys, "frozen", False):
        command += ["-m", "docwen_bundle.cli_entry"]
    return run_job_worker([*command, "serve", "--stdio"], parent_pid=parent_pid)
