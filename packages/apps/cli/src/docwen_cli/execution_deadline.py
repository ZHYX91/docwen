"""Cooperative command-wide deadlines for CLI application operations."""

from __future__ import annotations

import contextlib
import threading
from typing import Any, Protocol

from docwen_core.models.request import ConversionRequest


class Deadline(Protocol):
    """Deadline behavior consumed by the execution workflow."""

    @property
    def timed_out(self) -> bool: ...

    def register(self, request: ConversionRequest, *, batch: bool = False) -> object: ...

    def release(self, request: ConversionRequest, reservation: object) -> None: ...

    def finish(self) -> None: ...


class CancellationController(Protocol):
    """Current application cancellation surface required by CLI deadlines."""

    def prepare_execution_cancellation(self, request: ConversionRequest, *, batch: bool = False) -> object: ...

    def release_execution_cancellation(self, task_id: str, reservation: object) -> None: ...

    def cancel(self, task_id: str) -> None: ...


class ExecutionDeadline:
    """One command-wide cooperative deadline backed by application cancellation."""

    def __init__(self, controller: CancellationController, timeout_seconds: float) -> None:
        self._controller = controller
        self._timer = threading.Timer(timeout_seconds, self._expire)
        self._timer.daemon = True
        self._lock = threading.Lock()
        self._timed_out = threading.Event()
        self._reservations: dict[str, tuple[object, bool]] = {}
        self._finished = False

    @property
    def timed_out(self) -> bool:
        return self._timed_out.is_set()

    def start(self) -> ExecutionDeadline:
        self._timer.start()
        return self

    def register(self, request: ConversionRequest, *, batch: bool = False) -> object:
        reservation = self._controller.prepare_execution_cancellation(request, batch=batch)
        task_id = str(request.request_id)
        with self._lock:
            self._reservations[task_id] = (reservation, batch)
            cancel_now = self._timed_out.is_set()
        if cancel_now:
            self._cancel(task_id)
        return reservation

    def release(self, request: ConversionRequest, reservation: object) -> None:
        task_id = str(request.request_id)
        with self._lock:
            self._reservations.pop(task_id, None)
        self._controller.release_execution_cancellation(task_id, reservation)

    def finish(self) -> None:
        with self._lock:
            if self._finished:
                return
            self._finished = True
        self._timer.cancel()

    def close(self) -> None:
        self.finish()
        with self._lock:
            pending = list(self._reservations.items())
            self._reservations.clear()
        for task_id, (reservation, _batch) in pending:
            with contextlib.suppress(Exception):
                self._controller.release_execution_cancellation(task_id, reservation)

    def _expire(self) -> None:
        self._timed_out.set()
        with self._lock:
            task_ids = list(self._reservations)
        for task_id in task_ids:
            self._cancel(task_id)

    def _cancel(self, task_id: str) -> None:
        with contextlib.suppress(Exception):
            self._controller.cancel(task_id)


class NoopDeadline:
    """Deadline used by direct execution tests that do not own a command timer."""

    timed_out = False

    @staticmethod
    def register(request: ConversionRequest, *, batch: bool = False) -> object:
        del request, batch
        return object()

    @staticmethod
    def release(request: ConversionRequest, reservation: object) -> None:
        del request, reservation
        return None

    @staticmethod
    def finish() -> None:
        return None


def timeout_result(result: Any, *, timed_out: bool) -> Any:
    """Convert a successful or cancelled late result into a timeout result."""

    if not timed_out:
        return result
    error_type = str(getattr(getattr(result, "error", None), "error_type", ""))
    if not getattr(result, "success", False) and error_type not in {"cancelled", "operation_cancelled", "timeout"}:
        return result

    from docwen_core.models.result import ConversionErrorInfo, ConversionMetrics, ConversionResult

    return ConversionResult(
        task_id=str(getattr(result, "task_id", "timeout")),
        success=False,
        diagnostics=list(getattr(result, "diagnostics", [])),
        error=ConversionErrorInfo(
            error_type="operation_timeout",
            message="The DocWen operation exceeded its total timeout.",
            recoverable=True,
            diagnostic_code="operation_timeout",
        ),
        metrics=getattr(result, "metrics", ConversionMetrics()),
    )


__all__ = ["Deadline", "ExecutionDeadline", "NoopDeadline", "timeout_result"]
