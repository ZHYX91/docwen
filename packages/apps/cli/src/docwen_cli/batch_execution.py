"""Own batch workers and cancellation before waiting for their completion."""

from __future__ import annotations

import contextlib
import threading
import time
from collections.abc import Callable
from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from dataclasses import dataclass, field
from typing import Any

from docwen_cli.execution_deadline import Deadline, timeout_result
from docwen_core.models.request import ConversionRequest
from docwen_core.models.result import ConversionErrorInfo, ConversionMetrics, ConversionResult

_POLL_SECONDS = 0.05
_CANCELLATION_WAIT_SECONDS = 5.0


@dataclass(frozen=True)
class BatchExecution:
    results: dict[int, Any]
    interrupted: bool
    timed_out: bool


@dataclass
class _BatchWork:
    """Publication facts survive interruption inside executor.submit itself."""

    index: int
    request: ConversionRequest
    future: Future[Any] = field(default_factory=Future)
    reservation: object | None = None
    released: bool = False
    release_lock: Any = field(default_factory=threading.Lock, repr=False)

    def capture(self, reservation: object) -> None:
        if self.reservation is not None and self.reservation is not reservation:
            raise RuntimeError("Batch work received a different cancellation reservation")
        self.reservation = reservation

    def release(self, deadline: Deadline) -> None:
        with self.release_lock:
            if self.released or self.reservation is None:
                return
            deadline.release(self.request, self.reservation)
            self.released = True


def stopped_result(task_id: str, *, timed_out: bool, state: str) -> ConversionResult:
    """Keep unstarted work distinct from work whose cancellation is unconfirmed."""
    unconfirmed = state == "cancellation_unconfirmed"
    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type="operation_timeout" if timed_out else "operation_cancelled",
            message=(
                "Cancellation was requested, but the worker did not finish within the bounded wait."
                if unconfirmed
                else "Not started because the batch stopped before this input could execute."
            ),
            diagnostic_code="CLI_CANCELLATION_UNCONFIRMED" if unconfirmed else "CLI_NOT_STARTED",
            recoverable=not unconfirmed,
        ),
        metrics=ConversionMetrics(extra={"execution_state": state}),
    )


def execute_batch_requests(
    controller: Any,
    requests: list[tuple[int, ConversionRequest]],
    *,
    deadline: Deadline,
    max_workers: int,
    stop_on_error: bool,
    on_result: Callable[[int, Any], None] | None = None,
    on_submit: Callable[[int], None] | None = None,
) -> BatchExecution:
    """Stop submission and cancel owned requests before executor shutdown.

    Every request owns an independent Future before it can be submitted. A
    pool scheduling handle is never the only place holding the result: SIGINT
    may arrive after enqueue or publication but before submit returns.

    The cooperative drain is bounded. Running workers keep their cancelled
    Application reservation until their own finally block runs. This does not
    claim to forcibly terminate Python threads.
    """
    results: dict[int, Any] = {}
    works: dict[Future[Any], _BatchWork] = {}
    active: dict[Future[Any], _BatchWork] = {}
    executor = ThreadPoolExecutor(max_workers=max_workers)
    cursor = 0
    interrupted = False
    stop_submission = False

    def convert_one(work: _BatchWork) -> None:
        # Cancellation can win while a pool wrapper is still queued. Never
        # enter Application after its reservation has been released.
        if not work.future.set_running_or_notify_cancel():
            work.release(deadline)
            return
        try:
            result = controller.execute_single(work.request)
        except BaseException as exc:
            work.future.set_exception(exc)
        else:
            work.future.set_result(result)
        finally:
            work.release(deadline)

    def collect(completed: set[Future[Any]]) -> None:
        nonlocal stop_submission
        for future in completed:
            work = works[future]
            active.pop(future, None)
            if work.index in results:
                continue
            if future.cancelled():
                work.release(deadline)
                result = stopped_result(str(work.request.request_id), timed_out=deadline.timed_out, state="not_started")
            else:
                try:
                    result = future.result()
                except Exception as exc:
                    result = ConversionResult(
                        task_id=str(work.request.request_id),
                        success=False,
                        error=ConversionErrorInfo(error_type="conversion_failed", message=str(exc)),
                    )
                result = timeout_result(result, timed_out=deadline.timed_out)
            results[work.index] = result
            if stop_on_error and not result.success:
                stop_submission = True
            if on_result is not None:
                on_result(work.index, result)

    def cancel_owned() -> None:
        deadline.cancel_pending()
        for work in works.values():
            if work.future.cancel():
                work.release(deadline)

    try:
        try:
            while active or (cursor < len(requests) and not stop_submission):
                if deadline.timed_out:
                    break
                while len(active) < max_workers and cursor < len(requests) and not stop_submission:
                    if deadline.timed_out:
                        break
                    index, request = requests[cursor]
                    work = _BatchWork(index, request)
                    # Both the persistent record and the pending set exist
                    # before reserve/submit can expose an asynchronous window.
                    works[work.future] = work
                    active[work.future] = work
                    cursor += 1
                    deadline.register(request, on_reserved=work.capture)
                    executor.submit(convert_one, work)
                    if on_submit is not None:
                        on_submit(cursor)
                if not active:
                    break
                completed, _pending = wait(tuple(active), timeout=_POLL_SECONDS, return_when=FIRST_COMPLETED)
                collect(completed)
        except KeyboardInterrupt:
            interrupted = True
            deadline.interrupt()

        if interrupted or deadline.timed_out:
            cancel_owned()
            end = time.monotonic() + _CANCELLATION_WAIT_SECONDS
            while time.monotonic() < end:
                # A manually cancelled Future can be done before a queued
                # wrapper calls set_running_or_notify_cancel. Inspect done()
                # first instead of depending on the pool's notification.
                with contextlib.suppress(KeyboardInterrupt):
                    collect({future for future in active if future.done()})
                    pending = {future for future in active if not future.done()}
                    if not pending:
                        break
                    completed, _pending = wait(pending, timeout=_POLL_SECONDS, return_when=FIRST_COMPLETED)
                    collect(completed)
            # Recover any result committed even if SIGINT interrupted collect
            # after removing it from the active set but before presentation.
            collect({future for future in works if future.done()})
            for work in works.values():
                if work.index not in results:
                    results[work.index] = stopped_result(
                        str(work.request.request_id), timed_out=deadline.timed_out, state="cancellation_unconfirmed"
                    )
    except BaseException:
        deadline.interrupt()
        cancel_owned()
        raise
    finally:
        # The executor context manager would wait before cancellation handling.
        executor.shutdown(wait=False, cancel_futures=True)

    deadline.finish()
    for index, request in requests:
        if index not in results:
            results[index] = stopped_result(str(request.request_id), timed_out=deadline.timed_out, state="not_started")
    return BatchExecution(results, interrupted, deadline.timed_out)
