"""CLI 单元测试。"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

import docwen.cli.executor as executor
from docwen.services.batch import BatchResult, FileResult
from docwen.services.error_codes import ERROR_CODE_DEPENDENCY_MISSING, ERROR_CODE_INVALID_INPUT
from docwen.services.result import ConversionResult

pytestmark = pytest.mark.unit


def _load_golden(name: str) -> dict:
    path = Path(__file__).resolve().parents[1] / "fixtures" / "golden" / name
    return json.loads(path.read_text(encoding="utf-8"))


def test_execute_action_json_matches_golden_success(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], tmp_path: Path
) -> None:
    def _stub_execute(self, *_args, **_kwargs):
        return ConversionResult.ok(output_path="out", message="ok")

    monkeypatch.setattr(executor.ConversionService, "execute", _stub_execute, raising=True)

    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    exit_code = executor.execute_action("convert", str(f), json_mode=True)
    out = capsys.readouterr().out
    payload = json.loads(out)
    payload["data"]["input_file"] = "REPLACE_AT_RUNTIME"
    payload["data"]["duration"] = 0.0

    assert exit_code == 0
    assert payload == _load_golden("execute_action_success.json")


def test_execute_action_json_matches_golden_invalid_input(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], tmp_path: Path
) -> None:
    def _stub_execute(self, *_args, **_kwargs):
        return ConversionResult.fail(message="bad", error_code=ERROR_CODE_INVALID_INPUT, details="d")

    monkeypatch.setattr(executor.ConversionService, "execute", _stub_execute, raising=True)

    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    exit_code = executor.execute_action("convert", str(f), json_mode=True)
    out = capsys.readouterr().out
    payload = json.loads(out)
    payload["data"]["input_file"] = "REPLACE_AT_RUNTIME"
    payload["data"]["duration"] = 0.0

    assert exit_code != 0
    assert payload == _load_golden("execute_action_invalid_input.json")


def test_execute_batch_json_matches_golden_summary(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def _stub_execute_batch(self, batch_request, **_kwargs):
        return BatchResult(
            total=len(batch_request.requests),
            processed_count=len(batch_request.requests),
            results=[
                FileResult(file_path="a.docx", success=True, status="completed", output_path="out1", message="ok"),
                FileResult(
                    file_path="b.docx",
                    success=False,
                    status="failed",
                    message="nope",
                    error_code=ERROR_CODE_INVALID_INPUT,
                    details="d",
                ),
            ],
            success_count=1,
            failed_count=1,
            cancelled=False,
        )

    monkeypatch.setattr(executor.ConversionService, "execute_batch", _stub_execute_batch, raising=True)

    exit_code = executor.execute_batch(
        "convert",
        ["a.docx", "b.docx"],
        options={},
        json_mode=True,
        continue_on_error=True,
    )
    out = capsys.readouterr().out
    payload = json.loads(out)

    assert exit_code != 0
    assert payload == _load_golden("execute_batch_summary.json")


def test_execute_batch_json_matches_golden_summary_with_timing(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def _stub_execute_batch(self, batch_request, **_kwargs):
        return BatchResult(
            total=len(batch_request.requests),
            processed_count=len(batch_request.requests),
            results=[
                FileResult(
                    file_path="a.docx",
                    success=True,
                    status="completed",
                    output_path="out1",
                    message="ok",
                    duration_s=1.23,
                ),
                FileResult(
                    file_path="b.docx",
                    success=False,
                    status="failed",
                    message="nope",
                    error_code=ERROR_CODE_INVALID_INPUT,
                    details="d",
                    duration_s=0.45,
                ),
            ],
            success_count=1,
            failed_count=1,
            cancelled=False,
        )

    monkeypatch.setattr(executor.ConversionService, "execute_batch", _stub_execute_batch, raising=True)

    exit_code = executor.execute_batch(
        "convert",
        ["a.docx", "b.docx"],
        options={},
        json_mode=True,
        continue_on_error=True,
        include_timing=True,
    )
    out = capsys.readouterr().out
    payload = json.loads(out)

    assert exit_code != 0
    assert payload == _load_golden("execute_batch_summary_timing.json")


def test_execute_batch_json_matches_golden_cancelled(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def _stub_execute_batch(self, batch_request, **_kwargs):
        return BatchResult(
            total=len(batch_request.requests),
            processed_count=1,
            results=[
                FileResult(file_path="a.docx", success=True, status="completed", output_path="out1", message="ok"),
                FileResult(file_path="b.docx", success=False, status="cancelled", message="Cancelled by user"),
            ],
            success_count=1,
            failed_count=0,
            cancelled=True,
        )

    monkeypatch.setattr(executor.ConversionService, "execute_batch", _stub_execute_batch, raising=True)

    exit_code = executor.execute_batch(
        "convert",
        ["a.docx", "b.docx"],
        options={},
        json_mode=True,
        continue_on_error=True,
    )
    out = capsys.readouterr().out
    payload = json.loads(out)

    assert exit_code != 0
    assert payload == _load_golden("execute_batch_cancelled.json")


def test_execute_batch_json_matches_golden_stop_on_error(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def _stub_execute_batch(self, batch_request, **_kwargs):
        return BatchResult(
            total=len(batch_request.requests),
            processed_count=1,
            results=[
                FileResult(
                    file_path="a.docx",
                    success=False,
                    status="failed",
                    message="nope",
                    error_code=ERROR_CODE_INVALID_INPUT,
                    details="d",
                ),
            ],
            success_count=0,
            failed_count=1,
            cancelled=False,
        )

    monkeypatch.setattr(executor.ConversionService, "execute_batch", _stub_execute_batch, raising=True)

    exit_code = executor.execute_batch(
        "convert",
        ["a.docx", "b.docx"],
        options={},
        json_mode=True,
        continue_on_error=False,
    )
    out = capsys.readouterr().out
    payload = json.loads(out)

    assert exit_code != 0
    assert payload == _load_golden("execute_batch_stop_on_error.json")


def test_execute_batch_json_matches_golden_interrupted(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def _raise_interrupt(self, *_args, **_kwargs):
        raise KeyboardInterrupt()

    monkeypatch.setattr(executor.ConversionService, "execute_batch", _raise_interrupt, raising=True)

    exit_code = executor.execute_batch(
        "convert",
        ["a.docx", "b.docx"],
        options={},
        json_mode=True,
        continue_on_error=True,
    )
    out = capsys.readouterr().out
    payload = json.loads(out)

    assert exit_code == 130
    assert payload == _load_golden("execute_batch_interrupted.json")


def test_execute_batch_json_matches_golden_aggregate_failed(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def _stub_execute(self, *_args, **_kwargs):
        return ConversionResult.fail(message="nope", error_code=ERROR_CODE_DEPENDENCY_MISSING, details="d")

    monkeypatch.setattr(executor.ConversionService, "execute", _stub_execute, raising=True)

    exit_code = executor.execute_batch(
        "merge_pdfs",
        ["a.pdf", "b.pdf"],
        options={},
        json_mode=True,
        continue_on_error=True,
    )
    out = capsys.readouterr().out
    payload = json.loads(out)

    assert exit_code != 0
    assert payload == _load_golden("execute_batch_aggregate_failed.json")
