"""Identity-bound operations on a Windows process, never on a recycled PID."""

from __future__ import annotations

import ctypes
import subprocess
import sys
from ctypes import wintypes
from dataclasses import dataclass
from types import SimpleNamespace
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


def _windows_job_api() -> Any:
    """Keep the CPython handle-list launch and public Job APIs in one adapter."""
    if sys.platform != "win32":
        raise OSError("Windows Job ownership requires Windows")
    import _winapi
    import msvcrt

    import win32job
    import win32process

    # CPython's CreateProcess adapter consumes STARTUPINFO.lpAttributeList, as
    # Popen does, while retaining the initial thread handle needed to resume a
    # CREATE_SUSPENDED launch. Popen closes that handle before returning.
    return SimpleNamespace(
        kernel=_winapi,
        jobs=win32job,
        threads=win32process,
        os_handle=msvcrt.get_osfhandle,
        startup_info=subprocess.STARTUPINFO,
    )


class WindowsJobProcess:
    """A process starts suspended and cannot spawn before entering its Job.

    No breakaway flags are enabled. Kernel handles and Job membership, rather
    than process names or PID snapshots, establish signal and cleanup authority.
    """

    def __init__(self) -> None:
        self.api = _windows_job_api()
        self.pid = 0
        self.process: Any = None
        self.thread: Any = None
        self.assigned = False
        self.job = self.api.jobs.CreateJobObject(None, "")
        try:
            limits = self.api.jobs.QueryInformationJobObject(self.job, self.api.jobs.JobObjectExtendedLimitInformation)
            limits["BasicLimitInformation"]["LimitFlags"] |= self.api.jobs.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            self.api.jobs.SetInformationJobObject(self.job, self.api.jobs.JobObjectExtendedLimitInformation, limits)
        except BaseException:
            self.job.Close()
            raise

    def start(
        self,
        command: list[str],
        *,
        stdin: int,
        stdout: int,
        cwd: str | None,
        environment: dict[str, str] | None,
    ) -> None:
        kernel = self.api.kernel
        current = kernel.GetCurrentProcess()
        inherited: list[int] = []
        try:
            for stream in (stdin, stdout):
                inherited.append(
                    kernel.DuplicateHandle(
                        current, self.api.os_handle(stream), current, 0, True, kernel.DUPLICATE_SAME_ACCESS
                    )
                )
            startup = self.api.startup_info()
            startup.dwFlags |= 0x100  # STARTF_USESTDHANDLES
            startup.hStdInput, startup.hStdOutput = inherited
            startup.hStdError = startup.hStdOutput
            # Do not leak the owner's IPC or any unrelated inherited handle to
            # Office. Only the explicit file-backed standard streams cross.
            startup.lpAttributeList = {"handle_list": inherited}
            self.process, self.thread, self.pid, _thread_id = kernel.CreateProcess(
                None,
                subprocess.list2cmdline(command),
                None,
                None,
                True,
                0x4 | 0x08000000,  # CREATE_SUSPENDED | CREATE_NO_WINDOW
                environment,
                cwd,
                startup,
            )
            self.api.jobs.AssignProcessToJobObject(self.job, self.process)
            self.assigned = True
            if self.api.threads.ResumeThread(self.thread) != 1:
                raise OSError("Office primary thread did not have its expected suspended state")
        finally:
            for handle in inherited:
                kernel.CloseHandle(handle)

    def exited(self) -> bool:
        if self.process is None:
            return True
        if self.assigned:
            accounting = self.api.jobs.QueryInformationJobObject(
                self.job, self.api.jobs.JobObjectBasicAccountingInformation
            )
            if accounting["ActiveProcesses"]:
                return False
        state = self.api.kernel.WaitForSingleObject(self.process, 0)
        if state == 258:  # WAIT_TIMEOUT
            return False
        if state != 0:  # WAIT_OBJECT_0
            raise OSError("Cannot establish Office launcher handle exit")
        return True

    def terminate(self, *, force: bool) -> None:
        del force  # A headless Windows process has no safe generic soft signal.
        if self.process is None:
            return
        if self.assigned:
            self.api.jobs.TerminateJobObject(self.job, 1)
        elif self.api.kernel.WaitForSingleObject(self.process, 0) == 258:
            # Assignment failed before ResumeThread. No Office code, and hence
            # no Office child, has run; terminate only this retained handle.
            self.api.kernel.TerminateProcess(self.process, 1)

    def reap(self) -> int:
        if self.process is None:
            return -1
        if not self.exited():
            raise OSError("Office Job is still active")
        return int(self.api.kernel.GetExitCodeProcess(self.process))

    def close(self) -> None:
        try:
            self.job.Close()
        finally:
            try:
                if self.thread is not None:
                    self.api.kernel.CloseHandle(self.thread)
            finally:
                if self.process is not None:
                    self.api.kernel.CloseHandle(self.process)
