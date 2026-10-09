from __future__ import annotations

import ctypes
import os
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from tools import process_identity, run_lease, workspace_cleanup


@pytest.mark.integration
@pytest.mark.skipif(os.name != "nt", reason="Windows process handle semantics")
def test_windows_liveness_never_sends_a_console_signal(monkeypatch: pytest.MonkeyPatch) -> None:
    def forbidden_signal(*args: object) -> None:
        pytest.fail("A read-only liveness query must not call os.kill on Windows")

    monkeypatch.setattr(os, "kill", forbidden_signal)
    assert workspace_cleanup._process_alive(os.getpid())
    with subprocess.Popen([sys.executable, "-c", "pass"]) as process:
        process.wait(timeout=10)
        assert not workspace_cleanup._process_alive(process.pid)


@pytest.mark.unit
@pytest.mark.skipif(os.name != "nt", reason="Windows process handle semantics")
@pytest.mark.parametrize("error,expected", [(87, False), (5, True), (8, True)])
def test_windows_query_failure_preserves_uncertain_owners(
    monkeypatch: pytest.MonkeyPatch, error: int, expected: bool
) -> None:
    kernel = SimpleNamespace(
        OpenProcess=Mock(return_value=None),
        WaitForSingleObject=Mock(),
        GetProcessTimes=Mock(return_value=0),
        CloseHandle=Mock(),
    )
    monkeypatch.setattr(ctypes, "WinDLL", Mock(return_value=kernel))
    monkeypatch.setattr(ctypes, "get_last_error", lambda: error)
    monkeypatch.setattr(process_identity, "_cim_process_identity", lambda pid: None)
    assert process_identity._windows_observe(123).alive is expected
    kernel.WaitForSingleObject.assert_not_called()


@pytest.mark.unit
@pytest.mark.skipif(os.name != "nt", reason="Windows process handle semantics")
@pytest.mark.parametrize("wait_result,expected", [(0, False), (258, True), (0xFFFFFFFF, True)])
def test_windows_query_closes_handle_and_preserves_uncertain_waits(
    monkeypatch: pytest.MonkeyPatch, wait_result: int, expected: bool
) -> None:
    kernel = SimpleNamespace(
        OpenProcess=Mock(return_value=123),
        WaitForSingleObject=Mock(return_value=wait_result),
        GetProcessTimes=Mock(return_value=0),
        CloseHandle=Mock(),
    )
    monkeypatch.setattr(ctypes, "WinDLL", Mock(return_value=kernel))
    assert process_identity._windows_observe(42).alive is expected
    kernel.OpenProcess.assert_called_once_with(0x00101000, False, 42)
    kernel.CloseHandle.assert_called_once_with(123)


@pytest.mark.unit
@pytest.mark.parametrize(
    "alive,observed,expected,result",
    [
        (True, "windows-filetime:0000000000000001", "windows-filetime:0000000000000001", True),
        (True, "windows-filetime:0000000000000002", "windows-filetime:0000000000000001", False),
        (True, None, "windows-filetime:0000000000000001", True),
        (True, "windows-filetime:0000000000000002", "malformed", True),
        (True, "windows-filetime:0000000000000002", None, True),
        (False, None, "windows-filetime:0000000000000001", False),
        (True, "windows-cim-filetime:0000000000000100", "windows-filetime:0000000000000109", True),
        (True, "windows-cim-filetime:0000000000000100", "windows-filetime:00000000000000ff", True),
        (True, "windows-cim-filetime:0000000000000200", "windows-filetime:0000000000000100", False),
    ],
)
def test_lease_checks_process_birth_conservatively(
    monkeypatch: pytest.MonkeyPatch, alive: bool, observed: str | None, expected: str | None, result: bool
) -> None:
    monkeypatch.setattr(process_identity, "observe", lambda pid: process_identity.ProcessObservation(alive, observed))
    assert workspace_cleanup._lease_process_alive({"pid": 123, "processIdentity": expected}) is result


@pytest.mark.integration
@pytest.mark.skipif(os.name != "nt", reason="Windows process creation identity")
def test_new_lease_binds_actual_process_creation(tmp_path: Path) -> None:
    lease = run_lease.lease_payload(tmp_path, owner="docwen.test", kind="process-identity")
    assert lease["processIdentity"] == process_identity.observe(os.getpid()).identity
    assert workspace_cleanup._lease_process_alive(lease)
    lease["processIdentity"] = "windows-filetime:0000000000000000"
    assert not workspace_cleanup._lease_process_alive(lease)


@pytest.mark.unit
@pytest.mark.parametrize(
    "output,code,expected",
    [
        ('{"pid":42,"ticks":"256"}', 0, "windows-cim-filetime:0000000000000100"),
        ('{"pid":43,"ticks":"256"}', 0, None),
        ('{"pid":42,"ticks":"bad"}', 0, None),
        ('{"pid":42,"ticks":"256"}', 1, None),
        ("{}", 0, None),
        ("", 0, None),
    ],
)
def test_cim_identity_requires_matching_pid_and_valid_birth(
    monkeypatch: pytest.MonkeyPatch,
    output: str,
    code: int,
    expected: str | None,
) -> None:
    monkeypatch.setattr(subprocess, "run", Mock(return_value=SimpleNamespace(returncode=code, stdout=output)))
    assert process_identity._cim_process_identity(42) == expected


@pytest.mark.unit
def test_cim_query_timeout_does_not_authorize_cleanup(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(subprocess, "run", Mock(side_effect=subprocess.TimeoutExpired("powershell", 10)))
    assert process_identity._cim_process_identity(42) is None
