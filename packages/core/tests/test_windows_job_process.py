"""Portable API-boundary checks; these do not claim a native Windows run."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from docwen_core import windows_process

pytestmark = pytest.mark.unit


@pytest.fixture
def api(monkeypatch: pytest.MonkeyPatch) -> SimpleNamespace:
    kernel, jobs, threads = MagicMock(), MagicMock(), MagicMock()
    kernel.CreateProcess.return_value = (101, 102, 103, 104)
    kernel.DuplicateHandle.side_effect = [201, 202]
    kernel.WaitForSingleObject.return_value = 0
    kernel.GetExitCodeProcess.return_value = 0
    jobs.QueryInformationJobObject.side_effect = lambda _job, kind: (
        {"BasicLimitInformation": {"LimitFlags": 0}}
        if kind == jobs.JobObjectExtendedLimitInformation
        else {"ActiveProcesses": 0}
    )
    jobs.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000
    threads.ResumeThread.return_value = 1
    api = SimpleNamespace(
        kernel=kernel,
        jobs=jobs,
        threads=threads,
        os_handle=lambda value: value,
        startup_info=lambda: SimpleNamespace(dwFlags=0),
    )
    monkeypatch.setattr(windows_process, "_windows_job_api", lambda: api)
    return api


def _start(family: windows_process.WindowsJobProcess) -> None:
    family.start(["office.exe", "input with spaces.docx"], stdin=10, stdout=11, cwd=None, environment=None)


def test_suspended_job_assignment_precedes_business_execution_and_handles_are_restricted(api: SimpleNamespace) -> None:
    events: list[str] = []
    api.jobs.AssignProcessToJobObject.side_effect = lambda *args: events.append("assigned")
    api.threads.ResumeThread.side_effect = lambda *args: events.append("resumed") or 1
    family = windows_process.WindowsJobProcess()
    _start(family)
    assert events == ["assigned", "resumed"]
    args = api.kernel.CreateProcess.call_args.args
    assert args[1] == 'office.exe "input with spaces.docx"'
    assert args[5] & 0x4 and args[5] & 0x08000000
    assert args[8].lpAttributeList == {"handle_list": [201, 202]}
    assert (args[8].hStdInput, args[8].hStdOutput, args[8].hStdError) == (201, 202, 202)
    assert family.pid == 103 and family.exited() and family.reap() == 0
    family.close()
    assert [call.args[0] for call in api.kernel.CloseHandle.call_args_list] == [201, 202, 102, 101]


def test_parent_exit_never_substitutes_for_an_active_job(api: SimpleNamespace) -> None:
    family = windows_process.WindowsJobProcess()
    _start(family)
    api.jobs.QueryInformationJobObject.return_value = {"ActiveProcesses": 1}
    api.jobs.QueryInformationJobObject.side_effect = None
    assert not family.exited()
    family.terminate(force=True)
    api.jobs.TerminateJobObject.assert_called_once_with(family.job, 1)
    assert not family.exited()  # TerminateJobObject is a request, not proof.
    api.jobs.QueryInformationJobObject.return_value = {"ActiveProcesses": 0}
    assert family.exited()
    family.close()


def test_failed_job_assignment_never_resumes_or_guesses_children(api: SimpleNamespace) -> None:
    family = windows_process.WindowsJobProcess()
    api.jobs.AssignProcessToJobObject.side_effect = OSError("nested job unavailable")
    with pytest.raises(OSError, match="nested job"):
        _start(family)
    api.threads.ResumeThread.assert_not_called()
    api.kernel.WaitForSingleObject.return_value = 258
    assert not family.exited()
    family.terminate(force=True)
    api.kernel.TerminateProcess.assert_called_once_with(101, 1)
    api.jobs.TerminateJobObject.assert_not_called()
    api.kernel.WaitForSingleObject.return_value = 0
    assert family.exited()
    family.close()


@pytest.mark.parametrize("state", [258, 0xFFFFFFFF, 0x80])
def test_job_zero_requires_a_successful_launcher_handle_wait(api: SimpleNamespace, state: int) -> None:
    family = windows_process.WindowsJobProcess()
    _start(family)
    api.kernel.WaitForSingleObject.return_value = state
    if state == 258:
        assert not family.exited()
    else:
        with pytest.raises(OSError, match="handle exit"):
            family.exited()
    family.close()


def test_job_query_failure_grants_no_cleanup_permission(api: SimpleNamespace) -> None:
    family = windows_process.WindowsJobProcess()
    _start(family)
    api.jobs.QueryInformationJobObject.side_effect = OSError("query refused")
    with pytest.raises(OSError, match="query refused"):
        family.exited()
    family.close()
