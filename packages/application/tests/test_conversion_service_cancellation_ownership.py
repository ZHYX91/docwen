"""Real Application owners keep cancellation identity through admission races."""

from __future__ import annotations

import threading
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from docwen_application.controller import ApplicationController
from docwen_application.conversion_contracts import ConversionServiceError
from docwen_application.conversion_service import ConversionService
from docwen_core.models.result import ConversionErrorInfo, ConversionResult

from ._conversion_service_support import MARKDOWN_TO_XLSX_CAPABILITY_ID, _Committer, _Controller, _request

pytestmark = pytest.mark.contract


class _Runtime:
    is_available = True

    def __init__(self) -> None:
        self.executed: list[str] = []
        self.cancelled: list[str] = []
        self.block = False
        self.entered = threading.Event()
        self.finish = threading.Event()

    def describe_capabilities(self):
        return _Controller().describe_runtime_capabilities()

    def execute(self, request):
        self.executed.append(request.request_id)
        self.entered.set()
        if self.block:
            assert self.finish.wait(5), "test must release its controlled Runtime"
        cancelled = request.request_id in self.cancelled
        return ConversionResult(
            task_id=request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="cancelled" if cancelled else "conversion_failed", message="controlled Runtime completion"
            ),
        )

    def cancel(self, task_id: str) -> None:
        self.cancelled.append(task_id)

    def shutdown(self) -> None:
        self.finish.set()


@pytest.fixture
def admitted_service(tmp_path: Path):
    source = tmp_path / "input.md"
    source.write_text("# Cancellation ownership\n", encoding="utf-8")
    staging = tmp_path / "output"
    staging.mkdir()
    runtime = _Runtime()
    controller = ApplicationController(runtime_port=runtime)
    service = ConversionService(controller, _Committer())
    request = _request(source, staging, capability_id=MARKDOWN_TO_XLSX_CAPABILITY_ID)
    return runtime, controller, service, request


@pytest.mark.parametrize("phase", ["accepted", "running"])
@pytest.mark.parametrize("shared_controller", [False, True])
def test_rejected_duplicate_cannot_release_original_cancellation(admitted_service, phase, shared_controller):
    runtime, controller, service, request = admitted_service
    first = service.accept(service.plan(request).plan_id, "task.original")
    original_scope = controller._cancellation_scopes[first]
    competitor = ConversionService(controller, _Committer()) if shared_controller else service
    other = service.accept(service.plan(request).plan_id, "task.unrelated")
    other_scope = controller._cancellation_scopes[other]
    runtime.block = phase == "running"

    with ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(service.execute_accepted, first) if runtime.block else None
        try:
            if future is not None:
                assert runtime.entered.wait(5)
            with pytest.raises(ConversionServiceError) as error:
                competitor.accept(competitor.plan(request).plan_id, first)
            assert error.value.code == "task_id_in_use"
            assert controller._cancellation_scopes[first] is original_scope
            assert not original_scope.token.is_cancelled
            assert not other_scope.token.is_cancelled
            assert service.cancel(first) == "cancellation_requested"
            assert original_scope.token.is_cancelled
            assert controller._cancellation_scopes[other] is other_scope
        finally:
            runtime.finish.set()
        outcome = future.result(timeout=5) if future is not None else service.execute_accepted(first)

    assert outcome.state == "cancelled"
    assert runtime.executed == ([first] if phase == "running" else [])
    assert runtime.cancelled == ([first] if phase == "running" else [])
    assert service.cancel(other) == "cancellation_requested"
    assert service.execute_accepted(other).state == "cancelled"
    assert controller._cancellation_scopes == {}


def test_concurrent_accepts_leave_exactly_one_cancellable_owner(admitted_service):
    _runtime, controller, service, request = admitted_service
    plans = [service.plan(request).plan_id for _ in range(2)]
    start = threading.Barrier(2)

    def accept(plan_id):
        start.wait(timeout=5)
        try:
            return service.accept(plan_id, "task.concurrent")
        except ConversionServiceError as error:
            return error.code

    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(accept, plans))
    assert sorted(results) == ["task.concurrent", "task_id_in_use"]
    assert set(controller._cancellation_scopes) == {"task.concurrent"}
    assert service.cancel("task.concurrent") == "cancellation_requested"
    assert service.execute_accepted("task.concurrent").state == "cancelled"
    assert controller._cancellation_scopes == {}


@pytest.mark.parametrize("phase", ["capture", "published", "task-record"])
def test_interrupted_admission_releases_only_its_captured_handle(admitted_service, monkeypatch, phase):
    _runtime, controller, service, request = admitted_service
    other = service.accept(service.plan(request).plan_id, "task.other")
    other_scope = controller._cancellation_scopes[other]
    prepare = controller.prepare_execution_cancellation
    captured = []

    def interrupted_prepare(request, *, batch=False, on_reserved: Callable[[object], None] | None = None):
        def capture(handle):
            captured.append(handle)
            assert on_reserved is not None
            on_reserved(handle)
            if phase == "capture":
                raise KeyboardInterrupt

        handle = prepare(request, batch=batch, on_reserved=capture)
        if phase == "published":
            raise KeyboardInterrupt
        return handle

    class InterruptedTasks(dict):
        def __setitem__(self, key, value):
            super().__setitem__(key, value)
            raise KeyboardInterrupt

    with monkeypatch.context() as patch:
        patch.setattr(controller, "prepare_execution_cancellation", interrupted_prepare)
        if phase == "task-record":
            patch.setattr(service, "_tasks", InterruptedTasks(service._tasks))
        with pytest.raises(KeyboardInterrupt):
            service.accept(service.plan(request).plan_id, "task.retry")
        assert "task.retry" not in service._tasks
        assert controller._cancellation_scopes == {other: other_scope}

    retried = service.accept(service.plan(request).plan_id, "task.retry")
    current = controller._cancellation_scopes[retried]
    assert len(captured) == 1 and captured[0] is not current
    controller.cancel(retried, reservation=captured[0])
    controller.release_execution_cancellation(retried, captured[0])
    assert not current.token.is_cancelled
    service.cancel(retried)
    assert service.execute_accepted(retried).state == "cancelled"
    assert controller._cancellation_scopes == {other: other_scope}


def test_delayed_service_cancel_cannot_cancel_a_new_same_id_owner(admitted_service, monkeypatch):
    runtime, controller, service, request = admitted_service
    task_id = service.accept(service.plan(request).plan_id, "task.reused")
    original = controller._cancellation_scopes[task_id]
    cancel_entered = threading.Event()
    release_cancel = threading.Event()
    cancel = controller.cancel
    seen = []

    def delayed_cancel(task_id, *, reservation=None):
        seen.append(reservation)
        cancel_entered.set()
        assert release_cancel.wait(5)
        cancel(task_id, reservation=reservation)

    monkeypatch.setattr(controller, "cancel", delayed_cancel)
    with ThreadPoolExecutor(max_workers=1) as executor:
        delayed = executor.submit(service.cancel, task_id)
        try:
            assert cancel_entered.wait(5)
            assert service.execute_accepted(task_id).state == "failed"
            successor = ConversionService(controller, _Committer())
            successor.accept(successor.plan(request).plan_id, task_id)
            current = controller._cancellation_scopes[task_id]
            assert current is not original
        finally:
            release_cancel.set()
        assert delayed.result(timeout=5) == "cancellation_requested"

    assert seen == [original]
    assert not current.token.is_cancelled
    assert runtime.cancelled == []
    successor.cancel(task_id)
    assert successor.execute_accepted(task_id).state == "cancelled"


def test_invalid_controller_handle_is_rejected_and_the_captured_owner_is_released(admitted_service, monkeypatch):
    _runtime, controller, service, request = admitted_service
    prepare = controller.prepare_execution_cancellation

    def wrong_return(request, *, batch=False, on_reserved=None):
        prepare(request, batch=batch, on_reserved=on_reserved)
        return object()

    monkeypatch.setattr(controller, "prepare_execution_cancellation", wrong_return)
    with pytest.raises(ConversionServiceError) as error:
        service.accept(service.plan(request).plan_id, "task.invalid")
    assert error.value.code == "invalid_cancellation_reservation"
    assert controller._cancellation_scopes == {}
    assert "task.invalid" not in service._tasks
