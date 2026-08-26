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
        def prepare_execution_cancellation(request: ConversionRequest, *, batch: bool = False) -> object:
            del request
            assert batch is False
            return reservation

        @staticmethod
        def cancel(task_id: str) -> None:
            assert task_id == "deadline-task"
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


def test_timed_out_success_is_reported_as_timeout_not_success(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from docwen_cli.commands.convert import _execute_single
    from docwen_cli.exit_codes import ExitCode
    from docwen_core.models.result import ConversionResult

    source = tmp_path / "sample.md"
    source.write_text("hello", encoding="utf-8")

    class _Controller:
        @staticmethod
        def execute_single(request):
            return ConversionResult(task_id=request.request_id, success=True)

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
