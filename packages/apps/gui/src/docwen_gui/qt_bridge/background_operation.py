"""Serial cancellable work with completion delivered on the QObject's thread."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal, Slot

from docwen_core.cancellation import CancellationToken


class _CompletionSignal(QObject):
    finished = Signal()


class _OperationTask(QRunnable):
    def __init__(self, work: Callable[[CancellationToken], Any]) -> None:
        super().__init__()
        self.signals = _CompletionSignal()
        self.token = CancellationToken()
        self.work = work
        self.result: Any = None
        self.error: Exception | None = None

    def run(self) -> None:
        try:
            self.token.check()
            self.result = self.work(self.token)
        except Exception as exc:
            self.error = exc
        finally:
            self.signals.finished.emit()


class BackgroundOperation(QObject):
    """Run the latest submitted operation; discard superseded results."""

    busy_changed = Signal(bool)

    def __init__(self, parent: QObject) -> None:
        super().__init__(parent)
        self._task: _OperationTask | None = None
        self._completion: Callable[[Any, Exception | None], None] | None = None
        self._pending: tuple[Callable[[CancellationToken], Any], Callable[[Any, Exception | None], None]] | None = None

    @property
    def busy(self) -> bool:
        return self._task is not None or self._pending is not None

    def submit(
        self, work: Callable[[CancellationToken], Any], completed: Callable[[Any, Exception | None], None]
    ) -> None:
        self.cancel()
        self._pending = (work, completed)
        self._start_pending()
        self.busy_changed.emit(True)

    def cancel(self) -> None:
        self._pending = None
        if self._task is not None:
            self._task.token.cancel()

    def _start_pending(self) -> None:
        if self._task is not None or self._pending is None:
            return
        work, self._completion = self._pending
        self._pending = None
        self._task = _OperationTask(work)
        self._task.signals.finished.connect(self._finished)
        # The pool owns execution independently of a window's QObject lifetime.
        # Destroying a view disconnects its receiver without destroying a running thread.
        QThreadPool.globalInstance().start(self._task)

    @Slot()
    def _finished(self) -> None:
        task = self._task
        if task is None:
            return
        completed = self._completion
        self._task = None
        self._completion = None
        try:
            if not task.token.is_cancelled and completed is not None:
                completed(task.result, task.error)
        finally:
            self._start_pending()
            self.busy_changed.emit(self.busy)
