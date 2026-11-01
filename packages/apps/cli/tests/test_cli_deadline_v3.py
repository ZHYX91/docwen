"""Command-wide timeout and cancellation contract tests."""

from __future__ import annotations

import argparse
import json
import threading
from pathlib import Path
from typing import Any, cast

import pytest

from docwen_application.runtime_capability_catalog import RuntimeRoute

pytestmark = pytest.mark.contract


def _markdown_to_docx_route() -> RuntimeRoute:
    return RuntimeRoute(
        id="test:markdown:docx",
        operation="conversion",
        source="markdown",
        source_category="markdown",
        target="docx",
        action_name="",
        available=True,
        state="available",
        options=(),
    )


def test_execution_deadline_cancels_registered_application_operation() -> None:
    from docwen_cli.execution_deadline import ExecutionDeadline
    from docwen_core.models.request import ConversionRequest, FileRef

    cancelled = threading.Event()
    released: list[tuple[str, object]] = []
    reservation = object()

    class _Controller:
        @staticmethod
        def prepare_execution_cancellation(
            request: ConversionRequest, *, batch: bool = False, on_reserved=None
        ) -> object:
            del request
            assert batch is False
            if on_reserved is not None:
                on_reserved(reservation)
            return reservation

        @staticmethod
        def cancel(task_id: str, *, reservation: object | None = None) -> None:
            assert task_id == "deadline-task"
            assert reservation is not None
            cancelled.set()

        @staticmethod
        def release_execution_cancellation(task_id: str, reservation: object) -> None:
            released.append((task_id, reservation))

    request = ConversionRequest(
        request_id="deadline-task",
        input_refs=[FileRef(path="sample.md", format="md", category="text")],
        target_format="docx",
    )
    deadline = ExecutionDeadline(_Controller(), 0.02).start()
    token = deadline.register(request)
    try:
        assert cancelled.wait(1)
        assert deadline.timed_out is True
    finally:
        deadline.release(request, token)
        deadline.close()

    assert released == [("deadline-task", reservation)]


def test_timed_out_cancellation_is_reported_as_timeout(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.commands.convert import _execute_single
    from docwen_cli.exit_codes import ExitCode
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    source = tmp_path / "sample.md"
    source.write_text("hello", encoding="utf-8")

    class _Controller:
        @staticmethod
        def execute_single(request):
            return ConversionResult(
                task_id=request.request_id,
                success=False,
                error=ConversionErrorInfo(error_type="cancelled", message="Uncommitted work cancelled"),
            )

    class _ExpiredDeadline:
        timed_out = True

        @staticmethod
        def register(_request, *, batch=False):
            return object()

        @staticmethod
        def release(_request, _reservation) -> None:
            return None

        @staticmethod
        def finish() -> None:
            return None

    args = argparse.Namespace(
        json=True,
        timing=False,
        quiet=True,
        verbose=False,
        command_path="convert",
        output_path=str(tmp_path / "sample.docx"),
        overwrite=False,
    )
    exit_code = _execute_single(
        _Controller(),
        "",
        str(source),
        "docx",
        {},
        args,
        json_mode=True,
        deadline=cast(Any, _ExpiredDeadline()),
        route=_markdown_to_docx_route(),
    )

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code == int(ExitCode.TIMEOUT)
    assert captured.err == ""
    assert payload["success"] is False
    assert payload["error"]["category"] == "timeout"
    assert payload["error"]["code"] == "operation_timeout"


def test_deadline_cancels_application_operation_and_publishes_no_artifact(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from docwen_application.controller import ApplicationController
    from docwen_cli.commands.convert import _execute_single
    from docwen_cli.execution_deadline import ExecutionDeadline
    from docwen_cli.exit_codes import ExitCode
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    source = tmp_path / "source.md"
    output = tmp_path / "must-not-exist.docx"
    source.write_text("hello", encoding="utf-8")
    cancelled = threading.Event()

    class _Runtime:
        is_available = True

        @staticmethod
        def reserve_cancellation(task_id: str) -> None:
            del task_id
            return None

        @staticmethod
        def release_cancellation(task_id: str) -> None:
            del task_id
            return None

        @staticmethod
        def shutdown() -> None:
            return None

        @staticmethod
        def cancel(task_id: str) -> None:
            del task_id
            cancelled.set()

        @staticmethod
        def execute(request):
            assert cancelled.wait(1)
            return ConversionResult(
                task_id=request.request_id,
                success=False,
                error=ConversionErrorInfo(error_type="cancelled", message="cancelled"),
            )

    controller = ApplicationController(runtime_port=_Runtime())
    args = argparse.Namespace(
        json=True,
        timing=False,
        quiet=True,
        verbose=False,
        command_path="convert",
        output_path=str(output),
        overwrite=False,
    )
    deadline = ExecutionDeadline(controller, 0.02).start()
    try:
        exit_code = _execute_single(
            controller,
            "",
            str(source),
            "docx",
            {},
            args,
            json_mode=True,
            deadline=deadline,
            route=_markdown_to_docx_route(),
        )
    finally:
        deadline.close()

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == int(ExitCode.TIMEOUT)
    assert payload["error"]["code"] == "operation_timeout"
    # A cold pre-admission inspection may observe the deadline before Runtime
    # starts.  If Runtime does start, its assertion above proves the cancel was
    # forwarded.  Both valid races must produce the same public timeout result.
    assert not output.exists()


def test_expired_rejected_registration_does_not_cancel_or_release_another_owner(tmp_path, monkeypatch):
    from docwen_application.controller import ApplicationController, ControllerError
    from docwen_cli.execution_deadline import ExecutionDeadline
    from docwen_core.models.request import ConversionRequest, FileRef

    source = tmp_path / "source.csv"
    source.write_bytes(b"Name,Amount\nfirst,7\n")
    request = ConversionRequest(
        request_id="existing-owner",
        input_refs=[FileRef(path=str(source), format="csv", category="spreadsheet")],
        target_format="xlsx",
    )
    controller = ApplicationController()
    original = controller.prepare_execution_cancellation(request)
    original_prepare = controller.prepare_execution_cancellation
    deadline = ExecutionDeadline(controller, 0.02)

    def delayed_prepare(actual, *, batch=False, on_reserved=None):
        assert deadline._timed_out.wait(2)
        deadline._timer.join(timeout=2)
        assert not deadline._timer.is_alive()
        return original_prepare(actual, batch=batch, on_reserved=on_reserved)

    monkeypatch.setattr(controller, "prepare_execution_cancellation", delayed_prepare)
    deadline.start()
    try:
        with pytest.raises(ControllerError, match="already owned"):
            deadline.register(request)
        deadline.close()
        assert controller._cancellation_scopes[request.request_id] is original
        assert original.retain_until_release and not original.token.is_cancelled and not original.claimed
        assert deadline._reservations == {} and deadline._request_ids == set()
    finally:
        deadline.close()
        controller.release_execution_cancellation(request.request_id, original)
    assert controller._cancellation_scopes == {}


def test_deadline_during_capture_is_delivered_after_alias_publication(tmp_path):
    from docwen_application.controller import ApplicationController
    from docwen_cli.execution_deadline import ExecutionDeadline
    from docwen_core.models.request import ConversionRequest, FileRef

    source = tmp_path / "source.csv"
    source.write_bytes(b"Name,Amount\nfirst,7\n")
    request = ConversionRequest(
        request_id="capture-expiry",
        input_refs=[FileRef(path=str(source), format="csv", category="spreadsheet")],
        target_format="xlsx",
    )
    controller = ApplicationController()
    captured = []
    deadline = ExecutionDeadline(controller, 0.02).start()

    def capture(handle):
        captured.append(handle)
        assert controller._cancellation_scopes == {}
        # The actual timer is allowed to expire while Application holds its
        # alias-publication lock. No cancellation may re-enter that lock here.
        assert deadline._timed_out.wait(2)

    reservation = None
    try:
        reservation = deadline.register(request, on_reserved=capture)
        assert captured == [reservation]
        scope = controller._cancellation_scopes[request.request_id]
        assert scope is reservation and scope.token.is_cancelled
    finally:
        if reservation is not None:
            deadline.release(request, reservation)
        deadline.close()
    assert controller._cancellation_scopes == {}
