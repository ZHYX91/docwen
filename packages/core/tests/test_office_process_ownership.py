"""Office cleanup requires a positively identified, newly created COM server."""

import os
import sys
import time
from collections import deque
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from docwen_core import office_bridge, windows_process
from docwen_core.windows_process import WindowsProcessIdentity

pytestmark = pytest.mark.unit


class _Handle:
    """Portable stand-in for the explicitly accepted PyHANDLE type."""

    def __int__(self) -> int:
        return 123


def _pid_lookup(monkeypatch: pytest.MonkeyPatch, calls: list[object]) -> None:
    def window_process(hwnd: object) -> tuple[int, int]:
        calls.append(hwnd)
        return 10, 42

    monkeypatch.setitem(sys.modules, "pywintypes", SimpleNamespace(HANDLEType=_Handle))
    monkeypatch.setitem(sys.modules, "win32process", SimpleNamespace(GetWindowThreadProcessId=window_process))


@pytest.mark.parametrize("attribute", ["Hwnd", "HWND"])
def test_com_pid_accepts_known_handle_property_spellings(monkeypatch: pytest.MonkeyPatch, attribute: str) -> None:
    calls: list[object] = []
    _pid_lookup(monkeypatch, calls)
    assert office_bridge._get_com_app_pid(SimpleNamespace(**{attribute: 123})) == 42
    assert calls == [123]


def test_com_pid_ignores_method_wrapper_and_reads_powerpoint_hwnd(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[object] = []
    _pid_lookup(monkeypatch, calls)

    def wrong_case_method() -> None:
        pytest.fail("Ownership lookup must not call an ambiguous COM method")

    assert office_bridge._get_com_app_pid(SimpleNamespace(Hwnd=wrong_case_method, HWND=123)) == 42
    assert calls == [123]


def test_com_pid_accepts_integer_conversion_only_for_verified_pyhandle_type(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[object] = []
    _pid_lookup(monkeypatch, calls)
    handle = _Handle()
    assert office_bridge._get_com_app_pid(SimpleNamespace(HWND=handle)) == 42
    assert calls == [123]


@pytest.mark.parametrize("invalid", [True, False, 0, "123", object()])
def test_com_pid_rejects_nonhandles_without_native_lookup(monkeypatch: pytest.MonkeyPatch, invalid: object) -> None:
    calls: list[object] = []
    _pid_lookup(monkeypatch, calls)
    assert office_bridge._get_com_app_pid(SimpleNamespace(Hwnd=invalid)) is None
    assert calls == []


def test_com_pid_does_not_coerce_arbitrary_integer_like_objects(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[object] = []
    _pid_lookup(monkeypatch, calls)

    class Untrusted:
        def __int__(self) -> int:
            pytest.fail("Ownership lookup must not coerce an arbitrary COM object")

    assert office_bridge._get_com_app_pid(SimpleNamespace(Hwnd=Untrusted())) is None
    assert calls == []


def test_com_pid_tries_exact_hwnd_after_other_property_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[object] = []
    _pid_lookup(monkeypatch, calls)

    class PowerPoint:
        HWND = 123

        @property
        def Hwnd(self) -> None:
            raise AttributeError("case-sensitive COM map")

    assert office_bridge._get_com_app_pid(PowerPoint()) == 42
    assert calls == [123]


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
        app.Documents.Open.assert_called_once()
        probe_path = Path(app.Documents.Open.call_args.args[0])
        assert probe_path.parent == tmp_path
        assert probe_path.suffix == ".rtf"
        assert not probe_path.exists()
        app.Documents.Open.return_value.Close.assert_called_once_with(SaveChanges=False)
    else:
        assert app.mock_calls == []
    app.Quit.assert_not_called()
    assert "Visible" not in app.__dict__
    assert "AutomationSecurity" not in app.__dict__
    client.Dispatch.assert_not_called()


def test_word_identifies_its_private_document_window_before_changing_application(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    app = MagicMock()
    probe = MagicMock()
    document = MagicMock()
    window = probe.Windows.Item.return_value
    input_path = tmp_path / "in.doc"

    def opened(path: str, **kwargs: object) -> MagicMock:
        if path == str(input_path):
            assert app.AutomationSecurity == 3
            probe.Close.assert_called_once_with(SaveChanges=False)
            return document
        assert Path(path).read_bytes() == b"{\\rtf1\\ansi DocWen ownership probe.}"
        assert "AutomationSecurity" not in app.__dict__
        return probe

    app.Documents.Open.side_effect = opened
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
        assert "AutomationSecurity" not in app.__dict__
        document.SaveAs.assert_not_called()

    assert office_bridge._try_com_conversion(
        str(input_path),
        str(output),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
        on_process_owned=notified,
    ) == str(output)
    assert app.Documents.Open.call_count == 2
    probe.Windows.Item.assert_called_once_with(1)
    probe.Close.assert_called_once_with(SaveChanges=False)
    document.Close.assert_called_once_with(SaveChanges=False)
    app.Quit.assert_called_once()
    assert not list(tmp_path.glob(".dw-word-probe-*"))


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
