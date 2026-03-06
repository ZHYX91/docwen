"""services 单元测试。"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from docwen.services.batch import BatchEvent, execute_batch
from docwen.services.cancellation import CancellationToken
from docwen.services.error_codes import ERROR_CODE_INVALID_INPUT, ERROR_CODE_SKIPPED_SAME_FORMAT
from docwen.services.requests import BatchRequest, ConversionRequest
from docwen.services.result import ConversionResult

pytestmark = pytest.mark.unit


def test_batch_event_sink_emits_started_and_finished_with_result(tmp_path: Path) -> None:
    a = tmp_path / "a.txt"
    b = tmp_path / "b.txt"
    a.write_text("x", encoding="utf-8")
    b.write_text("y", encoding="utf-8")

    events: list[BatchEvent] = []

    def event_sink(e: BatchEvent) -> None:
        events.append(e)

    def execute_one(req: ConversionRequest) -> ConversionResult:
        time.sleep(0.01)
        return ConversionResult.ok(output_path=req.file_path)

    batch_request = BatchRequest(
        requests=[ConversionRequest(file_path=str(a)), ConversionRequest(file_path=str(b))],
        continue_on_error=True,
        max_workers=2,
    )

    result = execute_batch(batch_request, execute_one=execute_one, event_sink=event_sink)
    assert result.total == 2
    assert result.processed_count == 2

    for file_path in (str(a), str(b)):
        started_idx = next(i for i, e in enumerate(events) if e.type == "file_started" and e.file_path == file_path)
        finished_event = next(e for e in events if e.type == "file_finished" and e.file_path == file_path)
        finished_idx = next(i for i, e in enumerate(events) if e is finished_event)

        assert started_idx < finished_idx
        assert finished_event.result is not None
        assert finished_event.result.file_path == file_path


def test_batch_stops_on_first_error_when_continue_on_error_is_false(tmp_path: Path) -> None:
    a = tmp_path / "a.txt"
    b = tmp_path / "b.txt"
    a.write_text("x", encoding="utf-8")
    b.write_text("y", encoding="utf-8")

    events: list[BatchEvent] = []

    def event_sink(e: BatchEvent) -> None:
        events.append(e)

    def execute_one(req: ConversionRequest) -> ConversionResult:
        if req.file_path == str(a):
            return ConversionResult.fail(message="nope", error_code=ERROR_CODE_INVALID_INPUT, details="d")
        return ConversionResult.ok(output_path=req.file_path)

    batch_request = BatchRequest(
        requests=[ConversionRequest(file_path=str(a)), ConversionRequest(file_path=str(b))],
        continue_on_error=False,
        max_workers=4,
    )

    result = execute_batch(batch_request, execute_one=execute_one, event_sink=event_sink)
    assert result.total == 2
    assert result.processed_count == 1
    assert [r.file_path for r in result.results] == [str(a)]
    assert len([e for e in events if e.file_path == str(b)]) == 0


def test_batch_maps_skipped_same_format_to_skipped_status(tmp_path: Path) -> None:
    a = tmp_path / "a.txt"
    a.write_text("x", encoding="utf-8")

    def execute_one(_req: ConversionRequest) -> ConversionResult:
        return ConversionResult.ok(message="Already DOCX", error_code=ERROR_CODE_SKIPPED_SAME_FORMAT)

    batch_request = BatchRequest(
        requests=[ConversionRequest(file_path=str(a))],
        continue_on_error=True,
        max_workers=1,
    )

    result = execute_batch(batch_request, execute_one=execute_one)
    assert result.total == 1
    assert result.results[0].success is True
    assert result.results[0].status == "skipped"


def test_batch_cancel_marks_remaining_requests_as_cancelled(tmp_path: Path) -> None:
    a = tmp_path / "a.txt"
    b = tmp_path / "b.txt"
    a.write_text("x", encoding="utf-8")
    b.write_text("y", encoding="utf-8")

    cancel_token = CancellationToken()
    events: list[BatchEvent] = []

    def event_sink(e: BatchEvent) -> None:
        events.append(e)

    def execute_one(req: ConversionRequest) -> ConversionResult:
        if req.file_path == str(a):
            cancel_token.cancel()
            return ConversionResult.ok(output_path=req.file_path)
        return ConversionResult.ok(output_path=req.file_path)

    batch_request = BatchRequest(
        requests=[ConversionRequest(file_path=str(a)), ConversionRequest(file_path=str(b))],
        continue_on_error=True,
        max_workers=1,
    )

    result = execute_batch(batch_request, execute_one=execute_one, cancel_token=cancel_token, event_sink=event_sink)
    assert result.total == 2
    assert result.cancelled is True
    assert [r.file_path for r in result.results] == [str(a), str(b)]
    assert result.results[1].status == "cancelled"
    assert len([e for e in events if e.file_path == str(b)]) == 0
