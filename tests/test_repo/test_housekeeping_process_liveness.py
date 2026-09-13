from __future__ import annotations

import ctypes
import os
import subprocess
import sys
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from tools import workspace_cleanup


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
    kernel = SimpleNamespace(OpenProcess=Mock(return_value=None), WaitForSingleObject=Mock(), CloseHandle=Mock())
    monkeypatch.setattr(ctypes, "WinDLL", Mock(return_value=kernel))
    monkeypatch.setattr(ctypes, "get_last_error", lambda: error)
    assert workspace_cleanup._windows_process_alive(123) is expected
    kernel.WaitForSingleObject.assert_not_called()


@pytest.mark.unit
@pytest.mark.skipif(os.name != "nt", reason="Windows process handle semantics")
@pytest.mark.parametrize("wait_result,expected", [(0, False), (258, True), (0xFFFFFFFF, True)])
def test_windows_query_closes_handle_and_preserves_uncertain_waits(
    monkeypatch: pytest.MonkeyPatch, wait_result: int, expected: bool
) -> None:
    kernel = SimpleNamespace(
        OpenProcess=Mock(return_value=123), WaitForSingleObject=Mock(return_value=wait_result), CloseHandle=Mock()
    )
    monkeypatch.setattr(ctypes, "WinDLL", Mock(return_value=kernel))
    assert workspace_cleanup._windows_process_alive(42) is expected
    kernel.OpenProcess.assert_called_once_with(0x00100000, False, 42)
    kernel.CloseHandle.assert_called_once_with(123)
