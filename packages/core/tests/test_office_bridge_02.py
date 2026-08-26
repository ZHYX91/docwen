"""Focused tests split from test_office_bridge.py."""

from __future__ import annotations

from ._office_bridge_support import (
    Path,
    pytest,
)

pytestmark = pytest.mark.unit


def test_try_com_conversion_can_suppress_converter_created_word_revisions(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A route-owned private copy may disable recording SaveAs normalization."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.rtf"
    input_path.write_bytes(b"docx")

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None:
            return None

        @staticmethod
        def CoUninitialize() -> None:
            return None

    class _Document:
        TrackRevisions = True

        def SaveAs(self, output: str, *, FileFormat: int) -> None:
            assert FileFormat == 6
            assert self.TrackRevisions is False
            Path(output).write_bytes(b"{\\rtf1 output}")

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False

    class _Documents:
        document = _Document()

        def Open(self, input_file: str, **kwargs: object) -> _Document:
            assert input_file == str(input_path.resolve())
            assert kwargs["ReadOnly"] is True
            return self.document

    class _WordApp:
        def __init__(self) -> None:
            self.Documents = _Documents()

        def Quit(self) -> None:
            return None

    class _Win32Client:
        @staticmethod
        def DispatchEx(prog_id: str) -> _WordApp:
            assert prog_id == "Word.Application"
            return _WordApp()

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="Word.Application",
        save_format=6,
        app_type="word",
        suppress_new_revisions=True,
    )

    assert result == str(output_path.resolve())


def test_try_com_conversion_terminates_dispatch_ex_process_when_quit_leaves_it_running(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Only an isolated DispatchEx process may be force-terminated after Quit."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.odt"
    input_path.write_bytes(b"docx")
    terminated: list[int] = []

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None:
            return None

        @staticmethod
        def CoUninitialize() -> None:
            return None

    class _Document:
        def SaveAs(self, output: str, *, FileFormat: int) -> None:
            assert FileFormat == 23
            Path(output).write_bytes(b"odt")

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False

    class _Documents:
        def Open(self, input_file: str, **kwargs: object) -> _Document:
            assert input_file == str(input_path.resolve())
            return _Document()

    class _WordApp:
        Hwnd = 12345

        def __init__(self) -> None:
            self.Documents = _Documents()

        def Quit(self) -> None:
            return None

    class _Win32Client:
        @staticmethod
        def DispatchEx(prog_id: str) -> _WordApp:
            assert prog_id == "Word.Application"
            return _WordApp()

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))
    monkeypatch.setattr(office_bridge, "_get_com_app_pid", lambda app: 4242)
    monkeypatch.setattr(office_bridge, "_wait_for_process_exit", lambda process_id: False)
    monkeypatch.setattr(office_bridge, "_terminate_process", lambda process_id: terminated.append(process_id))

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="Word.Application",
        save_format=23,
        app_type="word",
    )

    assert result == str(output_path.resolve())
    assert terminated == [4242]


def test_try_com_conversion_terminates_new_dispatch_ex_process_when_pid_is_unavailable(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """When Word exposes no Hwnd/PID, clean only the process added by DispatchEx."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.odt"
    input_path.write_bytes(b"docx")
    snapshots = iter([{100}, {100, 200}, {100}])
    terminated: list[int] = []

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None:
            return None

        @staticmethod
        def CoUninitialize() -> None:
            return None

    class _Document:
        def SaveAs(self, output: str, *, FileFormat: int) -> None:
            assert FileFormat == 23
            Path(output).write_bytes(b"odt")

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False

    class _Documents:
        def Open(self, input_file: str, **kwargs: object) -> _Document:
            assert input_file == str(input_path.resolve())
            return _Document()

    class _WordApp:
        def __init__(self) -> None:
            self.Documents = _Documents()

        def Quit(self) -> None:
            return None

    class _Win32Client:
        @staticmethod
        def DispatchEx(prog_id: str) -> _WordApp:
            assert prog_id == "Word.Application"
            return _WordApp()

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))
    monkeypatch.setattr(office_bridge, "_get_com_app_pid", lambda app: None)
    monkeypatch.setattr(office_bridge, "_snapshot_process_ids", lambda process_names: next(snapshots))
    monkeypatch.setattr(office_bridge, "_wait_for_process_exit", lambda process_id: False)
    monkeypatch.setattr(office_bridge, "_terminate_process", lambda process_id: terminated.append(process_id))

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="Word.Application",
        save_format=23,
        app_type="word",
    )

    assert result == str(output_path.resolve())
    assert terminated == [200]


def test_try_com_conversion_terminates_new_dispatch_fallback_process_when_pid_is_unavailable(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Dispatch fallback can also leave an isolated WPS/Office process behind."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.doc"
    input_path.write_bytes(b"docx")
    snapshots = iter([{300}, {300, 400}, {300}])
    terminated: list[int] = []

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None:
            return None

        @staticmethod
        def CoUninitialize() -> None:
            return None

    class _Document:
        def SaveAs(self, output: str, *, FileFormat: int) -> None:
            assert FileFormat == 0
            Path(output).write_bytes(b"doc")

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False

    class _Documents:
        def Open(self, input_file: str, **kwargs: object) -> _Document:
            assert input_file == str(input_path.resolve())
            return _Document()

    class _WordApp:
        def __init__(self) -> None:
            self.Documents = _Documents()

        def Quit(self) -> None:
            return None

    class _Win32Client:
        @staticmethod
        def DispatchEx(prog_id: str) -> _WordApp:
            assert prog_id == "Kwps.Application"
            raise RuntimeError("WPS does not support isolated DispatchEx")

        @staticmethod
        def Dispatch(prog_id: str) -> _WordApp:
            assert prog_id == "Kwps.Application"
            return _WordApp()

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))
    monkeypatch.setattr(office_bridge, "_get_com_app_pid", lambda app: None)
    monkeypatch.setattr(office_bridge, "_snapshot_process_ids", lambda process_names: next(snapshots))
    monkeypatch.setattr(office_bridge, "_wait_for_process_exit", lambda process_id: False)
    monkeypatch.setattr(office_bridge, "_terminate_process", lambda process_id: terminated.append(process_id))
    monkeypatch.setattr(office_bridge.time, "sleep", lambda seconds: None)

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="Kwps.Application",
        save_format=0,
        app_type="word",
    )

    assert result == str(output_path.resolve())
    assert terminated == [400]


def test_process_exists_uses_tasklist_on_windows(monkeypatch: pytest.MonkeyPatch) -> None:
    """Windows process checks must not use os.kill(pid, 0) as a probe."""
    from docwen_core import office_bridge

    class _Completed:
        stdout = '"WINWORD.EXE","4242","Console","1","12,345 K"\n'

    def fake_run(args: list[str], **kwargs: object) -> _Completed:
        assert args == ["tasklist", "/FI", "PID eq 4242", "/FO", "CSV", "/NH"]
        assert kwargs["check"] is False
        return _Completed()

    def fail_if_called(process_id: int, signal_value: int) -> None:
        raise AssertionError(f"os.kill should not probe Windows PID {process_id} with signal {signal_value}")

    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge.subprocess, "run", fake_run)
    monkeypatch.setattr(office_bridge.os, "kill", fail_if_called)

    assert office_bridge._process_exists(4242) is True


def test_bounded_com_conversion_returns_worker_result(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The supervisor should return a completed worker result without force cleanup."""
    from docwen_core import office_bridge

    output_path = tmp_path / "output.docx"

    class _Process:
        alive = True
        terminated = False
        killed = False

        def is_alive(self) -> bool:
            return self.alive

        def join(self, timeout: float | None = None) -> None:
            self.alive = False

        def terminate(self) -> None:
            self.terminated = True
            self.alive = False

        def kill(self) -> None:
            self.killed = True
            self.alive = False

    class _Connection:
        closed = False

        def poll(self, timeout: float | None = None) -> bool:
            return True

        def recv(self) -> str:
            return str(output_path.resolve())

        def close(self) -> None:
            self.closed = True

    process = _Process()
    connection = _Connection()
    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge, "_snapshot_process_ids", lambda process_names: {100})
    monkeypatch.setattr(
        office_bridge,
        "_start_com_conversion_process",
        lambda *args, **kwargs: (process, connection),
    )

    result = office_bridge._try_com_conversion_bounded(
        str(tmp_path / "input.xps"),
        str(output_path),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
        timeout_s=10,
    )

    assert result == str(output_path.resolve())
    assert process.terminated is False
    assert process.killed is False
    assert connection.closed is True


def test_bounded_com_conversion_stops_worker_and_cleans_partial_output_on_timeout(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A stuck COM call must not outlive its shared timeout or leave residue."""
    from docwen_core import office_bridge

    output_path = tmp_path / "partial.docx"
    output_path.write_bytes(b"partial")
    snapshots = iter([{100}, {100, 200}, {100}])
    terminated_office_processes: list[int] = []

    class _Process:
        alive = True
        terminated = False
        killed = False

        def is_alive(self) -> bool:
            return self.alive

        def join(self, timeout: float | None = None) -> None:
            return None

        def terminate(self) -> None:
            self.terminated = True
            self.alive = False

        def kill(self) -> None:
            self.killed = True
            self.alive = False

    class _Connection:
        closed = False

        def poll(self, timeout: float | None = None) -> bool:
            return False

        def close(self) -> None:
            self.closed = True

    process = _Process()
    connection = _Connection()
    monotonic_values = iter([10.0, 12.0])
    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge.time, "monotonic", lambda: next(monotonic_values))
    monkeypatch.setattr(office_bridge.time, "sleep", lambda seconds: None)
    monkeypatch.setattr(office_bridge, "_snapshot_process_ids", lambda process_names: next(snapshots))
    monkeypatch.setattr(office_bridge, "_terminate_process", terminated_office_processes.append)
    monkeypatch.setattr(
        office_bridge,
        "_start_com_conversion_process",
        lambda *args, **kwargs: (process, connection),
    )

    result = office_bridge._try_com_conversion_bounded(
        str(tmp_path / "input.xps"),
        str(output_path),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
        timeout_s=1,
    )

    assert result is None
    assert process.terminated is True
    assert process.killed is False
    assert connection.closed is True
    assert terminated_office_processes == [200]
    assert not output_path.exists()


def test_bounded_com_conversion_observes_parent_cancellation_without_pickling_token(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Parent-side cancellation should terminate COM without crossing the process boundary."""
    from docwen_core import office_bridge

    class _CancellationView:
        reads = 0

        @property
        def is_cancelled(self) -> bool:
            self.reads += 1
            return self.reads >= 2

    class _Process:
        alive = True
        terminated = False

        def is_alive(self) -> bool:
            return self.alive

        def join(self, timeout: float | None = None) -> None:
            return None

        def terminate(self) -> None:
            self.terminated = True
            self.alive = False

        def kill(self) -> None:
            self.alive = False

    class _Connection:
        def poll(self, timeout: float | None = None) -> bool:
            return False

        def close(self) -> None:
            return None

    process = _Process()
    captured_start_kwargs: dict[str, object] = {}

    def fake_start(*args: object, **kwargs: object) -> tuple[_Process, _Connection]:
        captured_start_kwargs.update(kwargs)
        return process, _Connection()

    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge, "_snapshot_process_ids", lambda process_names: set())
    monkeypatch.setattr(office_bridge, "_start_com_conversion_process", fake_start)

    result = office_bridge._try_com_conversion_bounded(
        str(tmp_path / "input.doc"),
        str(tmp_path / "output.docx"),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
        cancel=_CancellationView(),
        timeout_s=60,
    )

    assert result is None
    assert process.terminated is True
    assert "cancel" not in captured_start_kwargs
