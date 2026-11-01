"""Cooperative command-wide deadlines for CLI application operations."""

from __future__ import annotations

import contextlib
import threading
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, Protocol

from docwen_core.models.request import ConversionRequest


class Deadline(Protocol):
    """Deadline behavior consumed by the execution workflow."""

    @property
    def timed_out(self) -> bool: ...

    def register(
        self,
        request: ConversionRequest,
        *,
        batch: bool = False,
        on_reserved: Callable[[object], None] | None = None,
    ) -> object: ...

    def release(self, request: ConversionRequest, reservation: object) -> None: ...

    def finish(self) -> None: ...

    def interrupt(self) -> None: ...

    def cancel_pending(self) -> None: ...


class CancellationController(Protocol):
    """Current application cancellation surface required by CLI deadlines."""

    def prepare_execution_cancellation(
        self,
        request: ConversionRequest,
        *,
        batch: bool = False,
        on_reserved: Callable[[object], None] | None = None,
    ) -> object: ...

    def release_execution_cancellation(self, task_id: str, reservation: object) -> None: ...

    def cancel(self, task_id: str, *, reservation: object | None = None) -> None: ...


@dataclass(frozen=True)
class _OwnedReservation:
    request: ConversionRequest
    handle: object
    asynchronous_owner: bool


class ExecutionDeadline:
    """One command-wide cooperative deadline backed by application cancellation."""

    def __init__(self, controller: CancellationController, timeout_seconds: float) -> None:
        self._controller = controller
        self._timer = threading.Timer(timeout_seconds, self._expire)
        self._timer.daemon = True
        self._lock = threading.Lock()
        self._timed_out = threading.Event()
        self._interrupted = threading.Event()
        self._request_ids: set[str] = set()
        self._reservations: dict[str, _OwnedReservation] = {}
        self._finished = False

    @property
    def timed_out(self) -> bool:
        return self._timed_out.is_set()

    def start(self) -> ExecutionDeadline:
        self._timer.start()
        return self

    def register(
        self,
        request: ConversionRequest,
        *,
        batch: bool = False,
        on_reserved: Callable[[object], None] | None = None,
    ) -> object:
        task_id = str(request.request_id)
        with self._lock:
            if task_id in self._request_ids:
                raise ValueError(f"Cancellation request is already registered: {task_id!r}")
            self._request_ids.add(task_id)

        def capture(handle: object) -> None:
            # Application still holds its lock and has not published aliases.
            # Only record/hand off identity here; never call cancel under it.
            with self._lock:
                self._reservations[task_id] = _OwnedReservation(request, handle, on_reserved is not None)
            if on_reserved is not None:
                on_reserved(handle)

        try:
            reservation = self._controller.prepare_execution_cancellation(request, batch=batch, on_reserved=capture)
            with self._lock:
                owned = self._reservations.get(task_id)
                cancel_now = self._timed_out.is_set() or self._interrupted.is_set()
            if owned is None or owned.handle is not reservation:
                raise RuntimeError("Application did not capture the exact cancellation reservation")
            # Expiry can occur between capture and alias publication. Recheck
            # only after the Application lock has been released.
            if cancel_now:
                self._cancel(task_id)
            return reservation
        except BaseException:
            with self._lock:
                owned = self._reservations.get(task_id)
            if owned is not None:
                self.release(request, owned.handle)
            else:
                with self._lock:
                    self._request_ids.discard(task_id)
            raise

    def release(self, request: ConversionRequest, reservation: object) -> None:
        task_id = str(request.request_id)
        with self._lock:
            owned = self._reservations.get(task_id)
            if owned is None or owned.handle is not reservation:
                return
        self._controller.release_execution_cancellation(task_id, reservation)
        with self._lock:
            if self._reservations.get(task_id) is owned:
                self._reservations.pop(task_id)
                self._request_ids.discard(task_id)

    def finish(self) -> None:
        with self._lock:
            if self._finished:
                return
            self._finished = True
        self._timer.cancel()

    def interrupt(self) -> None:
        self._interrupted.set()
        self.finish()
        self.cancel_pending()

    def close(self) -> None:
        self.finish()
        self.cancel_pending()
        # Work callbacks own asynchronous reservations until confirmed exit.
        # Synchronous calls have already unwound before close; release their
        # captured handle even if SIGINT interrupted the caller's assignment.
        with self._lock:
            synchronous = [owned for owned in self._reservations.values() if not owned.asynchronous_owner]
        for owned in synchronous:
            self.release(owned.request, owned.handle)

    def cancel_pending(self) -> None:
        with self._lock:
            task_ids = list(self._reservations)
        for task_id in task_ids:
            self._cancel(task_id)

    def _expire(self) -> None:
        with self._lock:
            if self._finished:
                return
            self._timed_out.set()
        self.cancel_pending()

    def _cancel(self, task_id: str) -> None:
        with self._lock:
            owned = self._reservations.get(task_id)
        if owned is not None:
            with contextlib.suppress(Exception):
                self._controller.cancel(task_id, reservation=owned.handle)


class NoopDeadline:
    """Deadline used by direct execution tests that do not own a command timer."""

    timed_out = False

    @staticmethod
    def register(
        request: ConversionRequest,
        *,
        batch: bool = False,
        on_reserved: Callable[[object], None] | None = None,
    ) -> object:
        del request, batch
        reservation = object()
        if on_reserved is not None:
            on_reserved(reservation)
        return reservation

    @staticmethod
    def release(request: ConversionRequest, reservation: object) -> None:
        del request, reservation
        return None

    @staticmethod
    def finish() -> None:
        return None

    @staticmethod
    def interrupt() -> None:
        return None

    @staticmethod
    def cancel_pending() -> None:
        return None


def timeout_result(result: Any, *, timed_out: bool) -> Any:
    """Label deadline cancellation without rewriting completed publication.

    Application returns a successful result only after Runtime finalization.
    Its commit facts do not become false because the command timer expires
    while another item is running or while a committed result is returned.
    Runtime's cancellation checks reject genuinely late uncommitted output.
    """

    if not timed_out or getattr(result, "success", False):
        return result
    error_type = str(getattr(getattr(result, "error", None), "error_type", ""))
    if error_type not in {"cancelled", "operation_cancelled", "timeout"}:
        return result

    from docwen_core.models.result import ConversionErrorInfo, ConversionMetrics, ConversionResult

    return ConversionResult(
        task_id=str(getattr(result, "task_id", "timeout")),
        success=False,
        artifacts=list(getattr(result, "artifacts", [])),
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
