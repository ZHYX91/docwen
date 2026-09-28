"""Real queued delivery and stale-result rejection for interactive admission."""

import threading
from pathlib import Path

import pytest
from PySide6.QtCore import QThread, QTimer

from docwen_core.detection import inspect_file, inspect_utf8_markdown_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.gui


def test_destroyed_owner_does_not_destroy_running_worker(qtbot) -> None:
    from PySide6.QtCore import QCoreApplication, QEvent, QObject
    from shiboken6 import isValid

    from docwen_gui.qt_bridge.background_operation import BackgroundOperation

    owner = QObject()
    operation = BackgroundOperation(owner)
    entered = threading.Event()
    release = threading.Event()
    finished = threading.Event()
    completions = []

    def work(token):
        entered.set()
        release.wait(5)
        finished.set()

    try:
        operation.submit(work, lambda result, error: completions.append(result))
        qtbot.waitUntil(entered.is_set)
        owner.deleteLater()
        QCoreApplication.sendPostedEvents(owner, QEvent.Type.DeferredDelete)
        assert not isValid(owner)
        release.set()
        qtbot.waitUntil(finished.is_set)
        QCoreApplication.processEvents()
        assert completions == []
    finally:
        release.set()
        qtbot.waitUntil(finished.is_set)


def test_latest_input_wins_without_blocking_event_loop(qtbot, tmp_path: Path) -> None:
    first = tmp_path / "first.md"
    second = tmp_path / "second.md"
    first.write_text("first", encoding="utf-8")
    second.write_text("second", encoding="utf-8")
    entered = threading.Event()
    release = threading.Event()

    def inspect(path: str):
        if path == str(first):
            entered.set()
            release.wait(5)
        return inspect_file(path)

    vm = MainWindowViewModel(file_inspector=inspect)
    completions = []
    timer_fired = []
    try:
        vm.request_files([str(first)], lambda outcome: completions.append("stale"))
        qtbot.waitUntil(entered.is_set)
        QTimer.singleShot(0, lambda: timer_fired.append(True))
        qtbot.waitUntil(lambda: bool(timer_fired))
        vm.request_files([str(second)], lambda outcome: completions.append(QThread.currentThread()))
        release.set()
        qtbot.waitUntil(lambda: not vm.inspection_busy)
        assert [ref.path for ref in vm.files] == [str(second)]
        assert completions == [vm.thread()]
    finally:
        release.set()
        vm.cancel_inspection()
        qtbot.waitUntil(lambda: not vm.inspection_busy)


def test_clear_discards_inflight_inspection(qtbot, tmp_path: Path) -> None:
    source = tmp_path / "input.md"
    source.write_text("input", encoding="utf-8")
    entered = threading.Event()
    release = threading.Event()

    def inspect(path: str):
        entered.set()
        release.wait(5)
        return inspect_file(path)

    vm = MainWindowViewModel(file_inspector=inspect)
    try:
        vm.request_files([str(source)])
        qtbot.waitUntil(entered.is_set)
        vm.clear_files()
        release.set()
        qtbot.waitUntil(lambda: not vm.inspection_busy)
        assert vm.files == []
    finally:
        release.set()
        vm.cancel_inspection()
        qtbot.waitUntil(lambda: not vm.inspection_busy)


def test_batch_requests_accumulate_while_inspection_is_running(qtbot, tmp_path: Path) -> None:
    paths = [tmp_path / name for name in ("first.md", "second.md")]
    for path in paths:
        path.write_text(path.stem, encoding="utf-8")
    entered = threading.Event()
    release = threading.Event()

    def inspect(path: str):
        entered.set()
        release.wait(5)
        return inspect_file(path)

    vm = MainWindowViewModel(file_inspector=inspect)
    vm.set_mode("batch")
    try:
        vm.request_files([str(paths[0])])
        qtbot.waitUntil(entered.is_set)
        vm.request_files([str(paths[1])])
        release.set()
        qtbot.waitUntil(lambda: not vm.inspection_busy)
        assert [Path(ref.path) for ref in vm.files] == paths
    finally:
        release.set()
        vm.cancel_inspection()
        qtbot.waitUntil(lambda: not vm.inspection_busy)


@pytest.mark.parametrize("mode", ["single", "batch"])
def test_readding_path_refreshes_inspection(tmp_path: Path, mode: str) -> None:
    source = tmp_path / "input.md"
    source.write_text("first", encoding="utf-8")
    vm = MainWindowViewModel()
    vm.set_mode(mode)
    vm.add_files([str(source)])
    before = vm.files[0]
    source.write_text("updated content", encoding="utf-8")
    outcome = vm.add_files([str(source)])
    assert len(vm.files) == 1
    assert outcome.added == (vm.files[0],)
    assert vm.files[0].metadata != before.metadata


def test_reveal_fallback_runs_on_ui_thread(qtbot, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from PySide6.QtCore import QObject

    from docwen_gui import path_actions
    from docwen_gui.qt_bridge.background_operation import BackgroundOperation

    parent = QObject()
    operation = BackgroundOperation(parent)
    source = tmp_path / "input.md"
    source.write_text("input", encoding="utf-8")
    command_threads = []
    fallback_threads = []
    outcomes = []
    monkeypatch.setattr(path_actions, "_platform_key", lambda: "linux")

    def command(_command):
        command_threads.append(QThread.currentThread() == parent.thread())
        return path_actions.PathActionResult(False, error_code="command_failed")

    def fallback(_path):
        fallback_threads.append(QThread.currentThread())
        return path_actions.PathActionResult(True)

    monkeypatch.setattr(path_actions, "_run_command", command)
    monkeypatch.setattr(path_actions, "_open_with_desktop_services", fallback)
    try:
        path_actions.reveal_path_async(source, operation, outcomes.append)
        qtbot.waitUntil(lambda: not operation.busy)
        assert command_threads and not any(command_threads)
        assert fallback_threads == [parent.thread()]
        assert outcomes == [path_actions.PathActionResult(True, fallback_used=True)]
    finally:
        operation.cancel()
        qtbot.waitUntil(lambda: not operation.busy)


@pytest.mark.parametrize("synthetic_first", [False, True])
def test_batch_interleaving_preserves_each_paths_inspection_intent(
    qtbot,
    tmp_path: Path,
    synthetic_first: bool,
) -> None:
    ordinary = tmp_path / "ordinary.md"
    synthetic = tmp_path / "clipboard.md"
    ordinary.write_text("ordinary text", encoding="utf-8")
    synthetic.write_text("<html><body>literal clipboard text</body></html>", encoding="utf-8")
    first = synthetic if synthetic_first else ordinary
    second = ordinary if synthetic_first else synthetic
    entered = threading.Event()
    release = threading.Event()

    def first_inspector(path: str):
        entered.set()
        release.wait(5)
        return inspect_utf8_markdown_snapshot(path) if path == str(synthetic) else inspect_file(path)

    vm = MainWindowViewModel()
    vm.set_mode("batch")
    try:
        vm.request_files(
            [str(first)],
            file_inspector=first_inspector,
        )
        qtbot.waitUntil(entered.is_set)
        vm.request_files(
            [str(second)],
            file_inspector=inspect_utf8_markdown_snapshot if second == synthetic else inspect_file,
        )
        release.set()
        qtbot.waitUntil(lambda: not vm.inspection_busy)
        by_path = {Path(ref.path): ref for ref in vm.files}
        assert set(by_path) == {ordinary, synthetic}
        ordinary_fact = by_path[ordinary].metadata[FILE_INSPECTION_METADATA_KEY]
        synthetic_fact = by_path[synthetic].metadata[FILE_INSPECTION_METADATA_KEY]
        assert ordinary_fact["detection_method"] != "synthetic_markdown"
        assert synthetic_fact["detection_method"] == "synthetic_markdown"
        assert synthetic_fact["detected_format"] == "markdown"
    finally:
        release.set()
        vm.cancel_inspection()
        qtbot.waitUntil(lambda: not vm.inspection_busy)


def test_clipboard_inspection_lease_survives_supersede_and_clear_until_physical_end(
    qtbot,
    tmp_path: Path,
) -> None:
    from docwen_gui.clipboard_inputs import ClipboardInputStore

    store = ClipboardInputStore(tmp_path / "managed")
    first = store.create("# First\n", display_name_template="Clipboard {index}.md")
    second = store.create("# Second\n", display_name_template="Clipboard {index}.md")
    third = store.create("# Third\n", display_name_template="Clipboard {index}.md")
    first_entered = threading.Event()
    first_release = threading.Event()
    third_entered = threading.Event()
    third_release = threading.Event()

    def blocking_first(path: str):
        first_entered.set()
        first_release.wait(5)
        return inspect_utf8_markdown_snapshot(path)

    def reject_second(_path: str):
        raise ValueError("test rejection")

    def blocking_third(path: str):
        third_entered.set()
        third_release.wait(5)
        return inspect_utf8_markdown_snapshot(path)

    vm = MainWindowViewModel(
        retain_inspection_inputs=store.retain_inspection,
        release_inspection_inputs=store.release_inspection,
    )
    try:
        vm.request_files([first.path], file_inspector=blocking_first)
        qtbot.waitUntil(first_entered.is_set)
        vm.request_files([second.path], file_inspector=reject_second)
        assert Path(first.path).is_file()
        assert Path(second.path).is_file()

        first_release.set()
        qtbot.waitUntil(lambda: not vm.inspection_busy)
        assert not Path(first.path).exists()
        assert not Path(second.path).exists()
        assert vm.files == []

        vm.request_files([third.path], file_inspector=blocking_third)
        qtbot.waitUntil(third_entered.is_set)
        vm.clear_files()
        assert Path(third.path).is_file()
        third_release.set()
        qtbot.waitUntil(lambda: not vm.inspection_busy)
        assert not Path(third.path).exists()
    finally:
        first_release.set()
        third_release.set()
        vm.cancel_inspection()
        qtbot.waitUntil(lambda: not vm.inspection_busy)
        store.close()

