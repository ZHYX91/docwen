"""Own worker lifetime, input reservations and cancellation on the GUI thread."""

from __future__ import annotations

import contextlib
from collections.abc import Callable, Mapping
from types import MappingProxyType
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QObject, QThread, Signal, Slot

from docwen_gui.i18n import t as _t
from docwen_gui.qt_bridge.execution import ExecutionThread

if TYPE_CHECKING:
    from docwen_application.controller import ApplicationController
    from docwen_core.models.request import ConversionRequest
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel


class ExecutionSupervisor(QObject):
    """Keep each task's original owner until the worker emits ``finished``.

    Result projection is separate from releasing the native worker. A reported
    result, cancellation request or startup error never destroys a live thread.
    """

    result_ready = Signal(object, dict)
    failed = Signal(str, dict)
    warning = Signal(str)

    def __init__(self, view_model: MainWindowViewModel, parent: QObject) -> None:
        super().__init__(parent)
        self._view_model = view_model
        self._threads: dict[str, QThread] = {}
        self._owners: dict[QThread, tuple[str, ApplicationController, object]] = {}
        self._starting = False

    @property
    def threads(self) -> Mapping[str, QThread]:
        """Read-only live workers, retained until their queued cleanup runs."""
        return MappingProxyType(self._threads)

    @property
    def busy(self) -> bool:
        return self._starting or bool(self._threads)

    def launch(
        self,
        *,
        controller: ApplicationController,
        request: ConversionRequest,
        context: dict[str, Any],
        on_reserved: Callable[[], None],
        aggregate_action_name: str = "",
        batch_execution: bool = False,
    ) -> bool:
        if self.busy:
            self.warning.emit(
                _t(
                    "main_window.task_already_running",
                    "A task is already running. Cancel it before starting another one.",
                )
            )
            return False
        task_id = request.request_id
        reservation_missing = object()
        reservation: object = reservation_missing
        thread: ExecutionThread | None = None
        self._starting = True
        try:
            reservation = controller.prepare_execution_cancellation(request, batch=batch_execution)
            self._view_model.reserve_execution_inputs(
                task_id, tuple(context.get("file_paths") or [context.get("file_path", "")])
            )
            on_reserved()
            thread = ExecutionThread(
                controller=controller,
                request=request,
                context=context,
                aggregate_action_name=aggregate_action_name,
                batch_execution=batch_execution,
                parent=self,
            )
            thread.result_signal.connect(self.result_ready)
            thread.error_signal.connect(self.failed)
            thread.finished.connect(self._finished)
            self._threads[task_id] = thread
            self._owners[thread] = (task_id, controller, reservation)
            thread.start()
        except Exception as exc:
            if thread is not None and thread.isRunning():
                # A native worker may start before the binding reports failure.
                # Keep both reservations and the QObject owner until finished.
                with contextlib.suppress(Exception):
                    controller.cancel(task_id)
                self.warning.emit(
                    _t(
                        "main_window.thread_start_uncertain",
                        "The task worker started but startup reporting failed; cancellation was requested.",
                    )
                )
                return True
            self._threads.pop(task_id, None)
            self._view_model.release_execution_inputs(task_id)
            if thread is not None:
                self._owners.pop(thread, None)
            if reservation is not reservation_missing:
                with contextlib.suppress(Exception):
                    controller.release_execution_cancellation(task_id, reservation)
            if thread is not None:
                thread.deleteLater()
            self.failed.emit(str(exc), context)
            return False
        finally:
            self._starting = False
        return True

    def cancel(self, task_id: str) -> None:
        thread = self._threads[task_id]
        _task_id, controller, _reservation = self._owners[thread]
        controller.cancel(task_id)

    def cancel_all(self) -> None:
        for task_id in tuple(self._threads):
            try:
                self.cancel(task_id)
            except Exception as exc:
                self.warning.emit(
                    _t(
                        "main_window.close_cancel_failed",
                        "Could not request cancellation for {task_id}: {message}",
                        task_id=task_id,
                        message=str(exc),
                    )
                )

    @Slot()
    def _finished(self) -> None:
        thread = self.sender()
        if not isinstance(thread, QThread):
            return
        owner = self._owners.pop(thread, None)
        if owner is None:
            thread.deleteLater()
            return
        task_id, controller, reservation = owner
        try:
            controller.release_execution_cancellation(task_id, reservation)
        except Exception as exc:
            self.warning.emit(str(exc))
        finally:
            self._threads.pop(task_id, None)
            self._view_model.release_execution_inputs(task_id)
            thread.deleteLater()
