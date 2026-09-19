"""Execute frozen runtime requests off the GUI thread."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QObject, QThread, Signal

from docwen_gui.execution_admission import check_frozen_request

if TYPE_CHECKING:
    from docwen_application.controller import ApplicationController
    from docwen_core.models.request import ConversionRequest


class ExecutionThread(QThread):
    """Run a single conversion request off the UI thread.

    Uses direct QThread.run() override instead of moveToThread +
    thread.started.connect to avoid signal-delivery deadlock in the
    MainWindow context.
    """

    result_signal = Signal(object, dict)
    error_signal = Signal(str, dict)

    def __init__(
        self,
        *,
        controller: ApplicationController,
        request: ConversionRequest,
        context: dict[str, Any],
        aggregate_action_name: str = "",
        batch_execution: bool = False,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._controller = controller
        self._request = request
        self._context = context
        self._aggregate_action_name = aggregate_action_name
        self._batch_execution = batch_execution

    def run(self) -> None:
        try:
            check_frozen_request(self._request)
            if self._aggregate_action_name:
                result = self._controller.execute_aggregate(self._request, self._aggregate_action_name)
            elif self._batch_execution:
                result = self._controller.execute_batch(self._request)
            else:
                result = self._controller.execute_single(self._request)
            self.result_signal.emit(result, self._context)
        except Exception as exc:
            self.error_signal.emit(str(exc), self._context)
