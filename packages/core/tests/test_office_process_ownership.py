"""Office cleanup requires a positively identified, newly created COM server."""

import os
import sys
import time
from collections import deque
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from docwen_core import office_bridge, windows_process
from docwen_core.windows_process import WindowsProcessIdentity

pytestmark = pytest.mark.unit


@pytest.mark.skipif(sys.platform != "win32", reason="Windows process identity API")
def test_real_process_identity_rejects_preexisting_process() -> None:
    identity = WindowsProcessIdentity.capture(os.getpid(), started_after_ns=0)
    assert identity is not None
    assert identity.pid == os.getpid()
    assert WindowsProcessIdentity.capture(os.getpid(), started_after_ns=time.time_ns()) is None


@pytest.mark.parametrize("pid", [None, 100])
def test_com_without_ownership_does_not_change_or_quit_application(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, pid: int | None
) -> None:
    app = MagicMock()
    client = MagicMock()
    client.DispatchEx.return_value = app
    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (MagicMock(), client))
    monkeypatch.setattr(office_bridge, "_get_com_app_pid", lambda app: pid)
    monkeypatch.setattr(WindowsProcessIdentity, "capture", lambda *args, **kwargs: None)
    result = office_bridge._try_com_conversion(
        str(tmp_path / "in.doc"),
        str(tmp_path / "out.docx"),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
    )
    assert result is None
    if pid is None:
        app.Documents.Open.assert_called_once_with(
            str(tmp_path / "in.doc"), ReadOnly=True, ConfirmConversions=False, AddToRecentFiles=False
        )
        app.Documents.Open.return_value.Close.assert_called_once_with(SaveChanges=False)
    else:
        assert app.mock_calls == []
    app.Quit.assert_not_called()
    assert "Visible" not in app.__dict__
    client.Dispatch.assert_not_called()


def test_word_identifies_its_private_document_window_before_changing_application(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    app = MagicMock()
    document = app.Documents.Open.return_value
    window = document.Windows.Item.return_value
    client = MagicMock()
    client.DispatchEx.return_value = app
    identity = WindowsProcessIdentity(200, 1234)
    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (MagicMock(), client))
    monkeypatch.setattr(office_bridge, "_get_com_app_pid", lambda obj: 200 if obj is window else None)
    monkeypatch.setattr(WindowsProcessIdentity, "capture", lambda *args, **kwargs: identity)
    monkeypatch.setattr(WindowsProcessIdentity, "terminate_if_running", lambda *args, **kwargs: None)
    output = tmp_path / "out.docx"
    document.SaveAs.side_effect = lambda *args, **kwargs: output.write_bytes(b"converted")

    def notified(owned: WindowsProcessIdentity) -> None:
        assert owned is identity
        assert "Visible" not in app.__dict__
        document.SaveAs.assert_not_called()

    assert office_bridge._try_com_conversion(
        str(tmp_path / "in.doc"),
        str(output),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
        on_process_owned=notified,
    ) == str(output)
    app.Documents.Open.assert_called_once()
    document.Windows.Item.assert_called_once_with(1)
    document.Close.assert_called_once_with(SaveChanges=False)
    app.Quit.assert_called_once()


def test_com_does_not_fall_back_to_shared_dispatch(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    client = MagicMock()
    client.DispatchEx.side_effect = RuntimeError("isolated server unavailable")
    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (MagicMock(), client))
    assert (
        office_bridge._try_com_conversion(
            str(tmp_path / "in.doc"),
            str(tmp_path / "out.docx"),
            prog_id="Kwps.Application",
            save_format=16,
            app_type="word",
        )
        is None
    )
    client.Dispatch.assert_not_called()


@pytest.mark.parametrize("cancelled", [False, True])
def test_supervisor_cleans_only_reported_identity(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, cancelled: bool
) -> None:
    owned = WindowsProcessIdentity(200, 1234)
    messages = deque([owned] if cancelled else [owned, str(tmp_path / "out.docx")])
    connection = MagicMock()
    connection.poll.side_effect = lambda *args: bool(messages)
    connection.recv.side_effect = messages.popleft
    process = MagicMock()
    process.is_alive.return_value = True
    process.terminate.side_effect = lambda: setattr(process.is_alive, "return_value", False)
    process.join.side_effect = lambda **kwargs: setattr(process.is_alive, "return_value", False)
    cleanup = MagicMock()
    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge, "_start_com_conversion_process", lambda *args, **kwargs: (process, connection))
    monkeypatch.setattr(WindowsProcessIdentity, "terminate_if_running", lambda self, **kwargs: cleanup(self))
    # Cancellation occurs before the pending ownership message is consumed.
    cancellation = MagicMock(spec=["is_cancelled"])
    cancellation.is_cancelled.side_effect = [False, True] if cancelled else lambda: False
    output = tmp_path / "out.docx"
    output.write_bytes(b"partial")
    result = office_bridge._try_com_conversion_bounded(
        str(tmp_path / "in.doc"),
        str(output),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
        cancel=cancellation,
    )
    assert result == (None if cancelled else str(output))
    assert output.exists() is not cancelled
    cleanup.assert_called_once_with(owned)
    connection.close.assert_called_once()


@pytest.mark.parametrize(
    "actual_creation,wait_result,terminated", [(1234, 258, True), (5678, 258, False), (1234, 0, False)]
)
def test_cleanup_checks_creation_identity_on_same_handle(
    monkeypatch: pytest.MonkeyPatch, actual_creation: int, wait_result: int, terminated: bool
) -> None:
    library = MagicMock()
    library.OpenProcess.return_value = 99
    library.WaitForSingleObject.return_value = wait_result
    monkeypatch.setattr(windows_process.sys, "platform", "win32")
    monkeypatch.setattr(windows_process, "_kernel32", lambda: library)
    monkeypatch.setattr(windows_process, "_creation_time", lambda api, handle: actual_creation)
    WindowsProcessIdentity(200, 1234).terminate_if_running(grace_ms=50)
    assert library.TerminateProcess.called is terminated
    if terminated:
        library.TerminateProcess.assert_called_once_with(99, 1)
    library.CloseHandle.assert_called_once_with(99)


def test_capture_rejects_process_created_before_attempt(monkeypatch: pytest.MonkeyPatch) -> None:
    library = MagicMock()
    library.OpenProcess.return_value = 99
    monkeypatch.setattr(windows_process.sys, "platform", "win32")
    monkeypatch.setattr(windows_process, "_kernel32", lambda: library)
    monkeypatch.setattr(windows_process, "_creation_time", lambda api, handle: 116444736000000000)
    assert WindowsProcessIdentity.capture(100, started_after_ns=1000) is None
    library.CloseHandle.assert_called_once_with(99)
