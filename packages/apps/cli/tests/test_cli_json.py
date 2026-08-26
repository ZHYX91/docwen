"""Protocol 3 JSON presenter contracts."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

REQUIRED_FIELDS = {
    "protocol_version",
    "product_version",
    "success",
    "command",
    "data",
    "error",
    "warnings",
    "meta",
}


def test_success_envelope_exact_shape(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter

    JsonPresenter().present_data("resources list", {"resources": []})

    payload = json.loads(capsys.readouterr().out)
    assert set(payload) == REQUIRED_FIELDS
    assert payload["protocol_version"] == 3
    assert payload["product_version"] == "0.9.0"
    assert payload["success"] is True
    assert payload["error"] is None
    assert payload["meta"] == {}


def test_error_envelope_exact_shape(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter

    JsonPresenter().present_error(
        "convert",
        "target exists",
        error_code="output_exists",
        details={"path": "out.docx"},
        hint="Use --overwrite to replace it.",
    )

    payload = json.loads(capsys.readouterr().out)
    assert set(payload) == REQUIRED_FIELDS
    assert payload["success"] is False
    assert payload["data"] == {}
    assert payload["error"] == {
        "category": "conflict",
        "code": "output_exists",
        "message": "target exists",
        "details": {"path": "out.docx"},
        "hint": "Use --overwrite to replace it.",
    }


def test_failed_envelope_without_typed_error_is_rejected() -> None:
    from docwen_cli.protocol import make_envelope

    with pytest.raises(ValueError, match="typed error object"):
        make_envelope(command="convert", success=False)


def test_schema_requires_error_object_when_success_is_false() -> None:
    schema_path = Path(__file__).resolve().parents[4] / "docs" / "specs" / "json-contracts.schema.json"
    schema = json.loads(schema_path.read_text(encoding="utf-8"))

    failure_rule = next(rule for rule in schema["allOf"] if rule["if"]["properties"]["success"].get("const") is False)
    assert failure_rule["then"]["properties"]["error"] == {"type": "object"}


def test_warning_is_structured_and_cleared(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter

    presenter = JsonPresenter()
    presenter.add_warning("fallback used")
    presenter.present_data("doctor", {})
    first = json.loads(capsys.readouterr().out)
    presenter.present_data("doctor", {})
    second = json.loads(capsys.readouterr().out)

    assert first["warnings"] == [{"code": "warning", "message": "fallback used"}]
    assert second["warnings"] == []


def test_single_projects_only_warning_diagnostics(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_core.models.result import ConversionDiagnostic, ConversionResult

    result = ConversionResult(
        task_id="single-warning",
        success=True,
        diagnostics=[
            ConversionDiagnostic(level="warning", message="needs review", code="REVIEW"),
            ConversionDiagnostic(level="info", message="done", code="DONE"),
        ],
    )
    JsonPresenter().present_single(result, command="convert")

    payload = json.loads(capsys.readouterr().out)
    assert payload["warnings"] == [{"level": "warning", "code": "REVIEW", "message": "needs review", "location": ""}]


def test_batch_projects_result_warning_diagnostics_with_file_context(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_core.models.result import ConversionDiagnostic, ConversionResult

    result = ConversionResult(
        task_id="batch-warning",
        success=True,
        diagnostics=[ConversionDiagnostic(level="warning", message="fallback", code="FALLBACK")],
    )
    JsonPresenter().present_batch([result], command="batch convert", input_files=["中文 sample.docx"])

    payload = json.loads(capsys.readouterr().out)
    assert payload["warnings"] == [
        {
            "level": "warning",
            "code": "FALLBACK",
            "message": "fallback",
            "location": "",
            "file": "中文 sample.docx",
        }
    ]


def test_partial_batch_has_typed_top_level_error(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    results = [
        ConversionResult(task_id="ok", success=True),
        ConversionResult(
            task_id="bad",
            success=False,
            error=ConversionErrorInfo(error_type="invalid_input", message="bad input"),
        ),
    ]
    JsonPresenter().present_batch(results, command="batch convert")

    payload = json.loads(capsys.readouterr().out)
    assert payload["success"] is False
    assert payload["error"]["category"] == "partial_failure"
    assert payload["error"]["code"] == "batch_partial_failure"


def test_interrupted_batch_has_typed_top_level_error(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_core.models.result import ConversionResult

    JsonPresenter().present_batch(
        [ConversionResult(task_id="ok", success=True)],
        command="batch convert",
        interrupted=True,
    )

    payload = json.loads(capsys.readouterr().out)
    assert payload["success"] is False
    assert payload["error"]["category"] == "cancelled"
    assert payload["error"]["code"] == "operation_cancelled"


def test_internal_cancelled_result_is_normalized_to_public_operation_cancelled(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    result = ConversionResult(
        task_id="cancelled",
        success=False,
        error=ConversionErrorInfo(error_type="cancelled", message="Task was cancelled"),
    )
    JsonPresenter().present_single(result, command="convert")

    payload = json.loads(capsys.readouterr().out)
    assert payload["error"]["category"] == "cancelled"
    assert payload["error"]["code"] == "operation_cancelled"


def test_all_failed_batch_promotes_typed_item_error(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    results = [
        ConversionResult(
            task_id="bad-1",
            success=False,
            error=ConversionErrorInfo(error_type="unsupported_route", message="route 1 unavailable"),
        ),
        ConversionResult(
            task_id="bad-2",
            success=False,
            error=ConversionErrorInfo(error_type="invalid_input", message="bad input"),
        ),
    ]
    JsonPresenter().present_batch(results, command="batch convert")

    payload = json.loads(capsys.readouterr().out)
    assert payload["success"] is False
    assert payload["error"]["category"] == "unavailable"
    assert payload["error"]["code"] == "unsupported_route"
    assert payload["error"]["message"] == "route 1 unavailable"


def test_timing_is_milliseconds_in_meta(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_core.models.result import ConversionMetrics, ConversionResult

    result = ConversionResult(task_id="timed", success=True, metrics=ConversionMetrics(duration_ms=125.5))
    JsonPresenter(include_timing=True).present_single(result, command="convert")

    payload = json.loads(capsys.readouterr().out)
    assert payload["meta"] == {"duration_ms": 125.5}


def test_failed_result_uses_typed_error(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    result = ConversionResult(
        task_id="failed",
        success=False,
        error=ConversionErrorInfo(
            error_type="dependency_missing",
            message="LibreOffice missing",
            diagnostic_code="LO_NOT_FOUND",
        ),
    )
    JsonPresenter().present_single(result, command="convert")

    error = json.loads(capsys.readouterr().out)["error"]
    assert error["category"] == "dependency"
    assert error["code"] == "dependency_missing"
    assert error["details"] == "LO_NOT_FOUND"


def test_json_preserves_unicode(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter

    JsonPresenter().present_error("inspect", "文件不存在", error_code="file_not_found")

    assert json.loads(capsys.readouterr().out)["error"]["message"] == "文件不存在"
