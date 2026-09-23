"""Real-Qt ownership tests for execution startup and graceful window close."""

from __future__ import annotations

import threading
import time
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.gui


def test_queued_execution_close_is_cancelled_when_window_is_destroyed(qapp):
    """A drained worker may queue close just before the owner is deleted."""
    from PySide6.QtCore import QCoreApplication, QEvent

    from docwen_gui.main_window import MainWindow
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

    window = MainWindow(view_model=MainWindowViewModel(controller=None))
    window.setup_ui()
    window._execution_close_pending = True
    window._poll_execution_close()
    window.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    # The old context-free singleShot invoked QWidget.close on a dead wrapper.
    qapp.processEvents()


class _BlockingController:
    has_runtime = True
    config_port = None

    def __init__(self, release_event: threading.Event, *, release_on_cancel: bool) -> None:
        self.release_event = release_event
        self.release_on_cancel = release_on_cancel
        self.reservation = object()
        self.prepared: list[tuple[object, bool]] = []
        self.cancelled: list[str] = []
        self.released: list[tuple[str, object]] = []
        self.stop_count = 0

    def prepare_execution_cancellation(self, request: object, *, batch: bool = False) -> object:
        self.prepared.append((request, batch))
        return self.reservation

    def execute_single(self, request: object):
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        self.release_event.wait()
        return ConversionResult(
            task_id=str(getattr(request, "request_id", "")),
            success=False,
            error=ConversionErrorInfo(error_type="cancelled", message="Cancelled for shutdown"),
        )

    def cancel(self, task_id: str) -> None:
        self.cancelled.append(task_id)
        if self.release_on_cancel:
            self.release_event.set()

    def release_execution_cancellation(self, task_id: str, reservation: object) -> None:
        self.released.append((task_id, reservation))

    def stop(self) -> None:
        self.stop_count += 1


def _launch_blocking_execution(window, tmp_path, controller: _BlockingController, task_id: str):
    from docwen_gui.path_identity import normalize_path

    source = tmp_path / f"{task_id}.md"
    source.write_text("# lifecycle", encoding="utf-8")
    file_path = normalize_path(str(source))
    window._batch_list_vm.add_files([file_path])
    window._view_model._controller = controller
    request = SimpleNamespace(request_id=task_id, input_refs=[])
    context = {
        "request_id": task_id,
        "file_path": file_path,
        "file_paths": [file_path],
        "display_name": source.name,
        "total_count": 1,
    }

    def project_reserved_execution() -> None:
        window._view_model.begin_execution_telemetry(task_id, (task_id,))
        window._batch_list_vm.set_file_status(file_path, "processing", operation_id=task_id)
        window._action_area_vm.show_cancel()

    assert window._workflow.launch(
        controller=controller,
        request=request,
        context=context,
        project_reserved_execution=project_reserved_execution,
    )
    return request, window._execution.threads[task_id]


def test_close_drains_cooperative_worker_without_blocking_gui(
    main_window,
    qapp,
    qtbot,
    tmp_path,
) -> None:
    release_event = threading.Event()
    controller = _BlockingController(release_event, release_on_cancel=True)
    request, thread = _launch_blocking_execution(main_window, tmp_path, controller, "close-cooperative")
    main_window.show()
    qtbot.waitUntil(thread.isRunning, timeout=1000)

    started_at = time.monotonic()
    close_accepted = main_window.close()
    close_elapsed = time.monotonic() - started_at

    assert close_accepted is False
    assert close_elapsed < 0.25
    assert main_window._execution_close_pending is True
    assert main_window.isVisible() is True
    assert main_window.isEnabled() is False
    assert controller.cancelled == [request.request_id]
    assert controller.stop_count == 0

    qtbot.waitUntil(lambda: main_window._shutdown_finalized, timeout=3000)

    assert main_window._execution.threads == {}
    assert controller.released == [(request.request_id, controller.reservation)]
    assert controller.stop_count == 1
    assert main_window.isVisible() is False


def test_close_timeout_keeps_disabled_window_and_parented_worker_alive(
    main_window,
    qapp,
    qtbot,
    tmp_path,
) -> None:
    release_event = threading.Event()
    controller = _BlockingController(release_event, release_on_cancel=False)
    request, thread = _launch_blocking_execution(main_window, tmp_path, controller, "close-timeout")
    main_window._EXECUTION_DRAIN_TIMEOUT_SECONDS = 0.02
    main_window.show()
    qtbot.waitUntil(thread.isRunning, timeout=1000)

    try:
        assert main_window.close() is False
        qtbot.waitUntil(lambda: main_window._execution_drain_timed_out, timeout=1000)

        assert main_window._execution_close_pending is True
        assert main_window.isVisible() is True
        assert main_window.isEnabled() is False
        assert main_window._execution.threads == {request.request_id: thread}
        assert thread.isRunning() is True
        assert thread.parent() is main_window._execution
        assert controller.cancelled == [request.request_id]
        assert controller.stop_count == 0
        assert any("remain open" in row.message for row in main_window._info_area_vm.history_rows)
    finally:
        release_event.set()

    qtbot.waitUntil(lambda: main_window._shutdown_finalized, timeout=3000)
    assert main_window._execution.threads == {}
    assert controller.released == [(request.request_id, controller.reservation)]
    assert controller.stop_count == 1


def test_repeated_shutdown_intent_finalizes_controller_once(main_window, qapp, qtbot) -> None:
    release_event = threading.Event()
    controller = _BlockingController(release_event, release_on_cancel=False)
    main_window._view_model._controller = controller
    main_window.show()

    main_window._view_model.request_shutdown()
    qtbot.waitUntil(lambda: main_window._shutdown_finalized, timeout=1000)
    main_window._view_model.request_shutdown()
    qapp.processEvents()

    assert controller.stop_count == 1


@pytest.mark.parametrize("cancel_by_close", [False, True])
def test_cancellation_and_release_use_original_controller(main_window, qtbot, tmp_path, cancel_by_close):
    release_event = threading.Event()
    original = _BlockingController(release_event, release_on_cancel=True)
    request, thread = _launch_blocking_execution(main_window, tmp_path, original, "original-owner")
    replacement = _BlockingController(threading.Event(), release_on_cancel=False)
    main_window._view_model._controller = replacement
    qtbot.waitUntil(thread.isRunning)
    try:
        if cancel_by_close:
            assert not main_window.close()
        else:
            main_window._workflow.cancel()
        qtbot.waitUntil(lambda: not main_window._execution.busy, timeout=3000)
        assert original.cancelled == [request.request_id]
        assert original.released == [(request.request_id, original.reservation)]
        assert replacement.cancelled == []
        assert replacement.released == []
    finally:
        release_event.set()
        qtbot.waitUntil(lambda: not main_window._execution.busy, timeout=3000)


def test_supervisor_rejects_reentrant_start_from_projection(main_window, qtbot, tmp_path):
    from PySide6.QtCore import QThread

    released = threading.Event()
    controller = _BlockingController(released, release_on_cancel=True)
    source = tmp_path / "source.md"
    source.write_text("# Source", encoding="utf-8")
    request = SimpleNamespace(request_id="outer", input_refs=[])
    context = {"request_id": "outer", "file_path": str(source)}
    nested_results = []
    result_threads = []
    main_window._execution.result_ready.connect(lambda *_args: result_threads.append(QThread.currentThread()))

    def on_reserved():
        nested_results.append(
            main_window._execution.launch(
                controller=controller,
                request=SimpleNamespace(request_id="nested", input_refs=[]),
                context={"request_id": "nested", "file_path": str(source)},
                on_reserved=lambda: None,
            )
        )

    try:
        assert main_window._execution.launch(
            controller=controller, request=request, context=context, on_reserved=on_reserved
        )
        assert nested_results == [False]
        assert controller.prepared == [(request, False)]
        assert tuple(main_window._execution.threads) == ("outer",)
    finally:
        released.set()
        qtbot.waitUntil(lambda: not main_window._execution.busy, timeout=3000)
    assert result_threads == [main_window.thread()]
    assert controller.released == [("outer", controller.reservation)]


def test_close_during_admission_prevents_late_start(main_window, monkeypatch):
    controller = _BlockingController(threading.Event(), release_on_cancel=True)
    main_window._view_model._controller = controller
    request = SimpleNamespace(request_id="late", input_refs=[])
    context = {"request_id": "late", "file_path": "source.md", "display_name": "source.md"}
    monkeypatch.setattr(main_window._workflow, "_resolve_route", lambda **_kwargs: ("md", SimpleNamespace(options=())))
    monkeypatch.setattr(main_window._requests, "single", lambda **_kwargs: (request, context))

    def confirm(_request):
        assert main_window.close()
        return True

    monkeypatch.setattr(main_window, "_confirm_request_admission", confirm)
    main_window._workflow.single(file_path="source.md", target_format="md", action_name="", options={})
    assert main_window._shutdown_finalized
    assert not main_window._execution.busy
    assert controller.prepared == []
    assert controller.stop_count == 1


def test_release_failure_still_removes_finished_thread(main_window, qtbot, tmp_path) -> None:
    release_event = threading.Event()
    release_event.set()
    controller = _BlockingController(release_event, release_on_cancel=False)
    release_calls: list[tuple[str, object]] = []

    def fail_release(task_id: str, reservation: object) -> None:
        release_calls.append((task_id, reservation))
        raise RuntimeError("release failed")

    controller.release_execution_cancellation = fail_release  # type: ignore[method-assign]
    request, _thread = _launch_blocking_execution(main_window, tmp_path, controller, "release-failure")

    qtbot.waitUntil(lambda: request.request_id not in main_window._execution.threads, timeout=2000)

    assert release_calls == [(request.request_id, controller.reservation)]
    assert main_window._execution._owners == {}
    assert any(row.message == "release failed" for row in main_window._info_area_vm.history_rows)


def test_projection_failure_after_reservation_releases_and_rolls_back(main_window, tmp_path) -> None:
    from docwen_gui.path_identity import normalize_path

    release_event = threading.Event()
    controller = _BlockingController(release_event, release_on_cancel=False)
    task_id = "projection-failure"
    source = tmp_path / "projection-failure.md"
    source.write_text("# projection", encoding="utf-8")
    file_path = normalize_path(str(source))
    main_window._batch_list_vm.add_files([file_path])
    main_window._view_model._controller = controller
    request = SimpleNamespace(request_id=task_id)
    context = {
        "request_id": task_id,
        "file_path": file_path,
        "file_paths": [file_path],
        "display_name": source.name,
        "total_count": 1,
    }

    def fail_projection() -> None:
        main_window._view_model.begin_execution_telemetry(task_id, (task_id,))
        main_window._batch_list_vm.set_file_status(file_path, "processing", operation_id=task_id)
        main_window._action_area_vm.show_cancel()
        raise RuntimeError("projection failed")

    assert not main_window._workflow.launch(
        controller=controller,
        request=request,
        context=context,
        project_reserved_execution=fail_projection,
    )

    assert controller.prepared == [(request, False)]
    assert controller.released == [(task_id, controller.reservation)]
    assert main_window._execution.threads == {}
    assert main_window._execution._owners == {}
    assert main_window._action_area_vm.cancel_visible is False
    assert main_window._view_model._active_execution_id is None
    entry = main_window._batch_list_vm.get_file_entry(file_path)
    assert entry is not None
    assert entry.status == "failed"
    assert entry.error_message == "projection failed"


def test_thread_start_failure_before_running_releases_reservation(
    main_window,
    tmp_path,
    monkeypatch,
) -> None:
    from docwen_gui.path_identity import normalize_path
    from docwen_gui.qt_bridge.execution import ExecutionThread

    release_event = threading.Event()
    controller = _BlockingController(release_event, release_on_cancel=False)
    task_id = "start-failure"
    source = tmp_path / "start-failure.md"
    source.write_text("# start", encoding="utf-8")
    file_path = normalize_path(str(source))
    main_window._batch_list_vm.add_files([file_path])
    main_window._view_model._controller = controller
    request = SimpleNamespace(request_id=task_id)
    context = {
        "request_id": task_id,
        "file_path": file_path,
        "file_paths": [file_path],
        "display_name": source.name,
        "total_count": 1,
    }

    def fail_start(_thread: object) -> None:
        raise RuntimeError("start failed")

    monkeypatch.setattr(ExecutionThread, "start", fail_start)

    assert not main_window._workflow.launch(
        controller=controller,
        request=request,
        context=context,
        project_reserved_execution=lambda: main_window._batch_list_vm.set_file_status(
            file_path,
            "processing",
            operation_id=task_id,
        ),
    )

    assert controller.released == [(task_id, controller.reservation)]
    assert main_window._execution.threads == {}
    entry = main_window._batch_list_vm.get_file_entry(file_path)
    assert entry is not None
    assert entry.status == "failed"
    assert entry.error_message == "start failed"


def test_thread_start_error_after_native_start_retains_ownership_until_finished(
    main_window,
    qtbot,
    tmp_path,
    monkeypatch,
) -> None:
    from docwen_gui.qt_bridge.execution import ExecutionThread

    release_event = threading.Event()
    controller = _BlockingController(release_event, release_on_cancel=True)
    original_start = ExecutionThread.start

    def start_then_report_error(thread: ExecutionThread) -> None:
        original_start(thread)
        deadline = time.monotonic() + 1.0
        while not thread.isRunning() and time.monotonic() < deadline:
            time.sleep(0.001)
        assert thread.isRunning()
        raise RuntimeError("native start reporting failed")

    monkeypatch.setattr(ExecutionThread, "start", start_then_report_error)
    request, thread = _launch_blocking_execution(main_window, tmp_path, controller, "start-uncertain")

    assert main_window._execution.threads[request.request_id] is thread
    assert thread in main_window._execution._owners
    assert controller.cancelled == [request.request_id]
    assert controller.released == []

    qtbot.waitUntil(lambda: request.request_id not in main_window._execution.threads, timeout=2000)

    assert controller.released == [(request.request_id, controller.reservation)]
    assert main_window._execution._owners == {}
