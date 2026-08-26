from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace
from typing import Any, cast

import pytest

pytestmark = pytest.mark.gui


@pytest.mark.parametrize(
    ("attribute", "report_key"),
    [
        ("duration_ms", "durationMs"),
        ("input_bytes", "inputBytes"),
        ("output_bytes", "outputBytes"),
    ],
)
@pytest.mark.parametrize("value", [True, False])
def test_allowlisted_conversion_metrics_rejects_boolean_numeric_values(
    attribute: str,
    report_key: str,
    value: bool,
) -> None:
    from docwen_gui.release_smoke import _allowlisted_conversion_metrics

    metrics = SimpleNamespace(
        duration_ms=12.5,
        input_bytes=10,
        output_bytes=20,
        extra={"engine": "office_bridge", "backend": "fixture-office"},
    )
    setattr(metrics, attribute, value)
    expected: dict[str, object] = {
        "durationMs": 12.5,
        "inputBytes": 10,
        "outputBytes": 20,
        "engine": "office_bridge",
        "backend": "fixture-office",
    }
    expected[report_key] = None

    assert _allowlisted_conversion_metrics(SimpleNamespace(metrics=metrics)) == expected


@pytest.mark.parametrize("attribute", ["engine", "backend"])
@pytest.mark.parametrize("value", [None, True, 1, ["office_bridge"], "", "   "])
def test_allowlisted_conversion_metrics_rejects_invalid_identity_values(
    attribute: str,
    value: object,
) -> None:
    from docwen_gui.release_smoke import _allowlisted_conversion_metrics

    extra: dict[str, object] = {"engine": "office_bridge", "backend": "fixture-office"}
    extra[attribute] = value
    expected: dict[str, object] = {
        "durationMs": 12.5,
        "inputBytes": 10,
        "outputBytes": 20,
        "engine": "office_bridge",
        "backend": "fixture-office",
    }
    expected[attribute] = None
    metrics = SimpleNamespace(duration_ms=12.5, input_bytes=10, output_bytes=20, extra=extra)

    assert _allowlisted_conversion_metrics(SimpleNamespace(metrics=metrics)) == expected


@pytest.mark.parametrize("surface", ["panel", "action"])
def test_conversion_release_hook_uses_requested_main_window_surface(
    surface: str,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PySide6.QtCore import QTimer

    from docwen_gui.release_smoke import _schedule_test_conversion_report

    source = tmp_path / ("source.md" if surface == "action" else "source.docx")
    source.write_text("release smoke", encoding="utf-8")
    output = tmp_path / "outputs" / "source.pdf"
    output.parent.mkdir()
    output.write_bytes(b"%PDF-release-smoke")
    report = tmp_path / "report.json"

    env = {
        "DOCWEN_GUI_TEST_CONVERSION_REPORT": str(report),
        "DOCWEN_GUI_TEST_CONVERSION_INPUT": str(source),
        "DOCWEN_GUI_TEST_CONVERSION_OUTPUT_DIR": str(output.parent),
        "DOCWEN_GUI_TEST_CONVERSION_TARGET": "pdf",
        "DOCWEN_GUI_TEST_CONVERSION_SURFACE": surface,
    }
    for key, value in env.items():
        monkeypatch.setenv(key, value)

    config_batches: list[dict[str, object]] = []
    requests: list[tuple[str, str]] = []
    handler_calls: list[dict[str, object]] = []

    class _ConfigPort:
        def set_many(self, values: dict[str, object]) -> bool:
            config_batches.append(dict(values))
            return True

    class _ViewModel:
        controller = SimpleNamespace(config_port=_ConfigPort())

        def add_files(self, paths: list[str]) -> None:
            assert paths == [str(source)]

    entry = SimpleNamespace(status="processing", output_path=str(output), error_message="")
    result = SimpleNamespace(
        metrics=SimpleNamespace(
            duration_ms=12.5,
            input_bytes=source.stat().st_size,
            output_bytes=output.stat().st_size,
            extra={"engine": "office_bridge", "backend": "fixture-office", "ignored": "not evidence"},
        )
    )
    window_holder: dict[str, Any] = {}

    def finish_conversion(request_surface: str, target: str) -> None:
        requests.append((request_surface, target))
        window_holder["window"]._on_execution_finished(result, {"surface": request_surface})

    def original_execution_finished(observed_result: object, context: dict[str, object]) -> None:
        assert observed_result is result
        handler_calls.append(context)
        entry.status = "completed"

    window = SimpleNamespace(
        view_model=_ViewModel(),
        _action_area_vm=SimpleNamespace(
            visible=True,
            request_conversion=lambda target: finish_conversion("action", target),
        ),
        _conversion_panel_vm=SimpleNamespace(
            current_file_path=str(source),
            request_conversion=lambda target: finish_conversion("panel", target),
        ),
        _batch_list_vm=SimpleNamespace(get_file_entry=lambda _normalized: entry),
        _on_execution_finished=original_execution_finished,
        close=lambda: None,
    )
    window_holder["window"] = window
    app = SimpleNamespace(processEvents=lambda: None)
    monkeypatch.setattr(QTimer, "singleShot", lambda _delay, callback: callback())

    _schedule_test_conversion_report(cast(Any, app), cast(Any, window))

    assert requests == [(surface, "pdf")]
    assert handler_calls == [{"surface": surface}]
    assert config_batches == [
        {
            "output.directory.mode": "custom",
            "output.directory.custom_path": str(output.parent),
            "output.directory.create_date_subfolder": False,
        }
    ]
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["success"] is True
    assert payload["status"] == "completed"
    assert payload["surface"] == surface
    assert payload["targetFormat"] == "pdf"
    assert payload["outputPath"] == str(output)
    assert payload["outputBytes"] == output.stat().st_size
    assert payload["conversionMetrics"] == {
        "durationMs": 12.5,
        "inputBytes": source.stat().st_size,
        "outputBytes": output.stat().st_size,
        "engine": "office_bridge",
        "backend": "fixture-office",
    }


def test_conversion_release_hook_captures_rendered_successful_warning(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PySide6.QtCore import QTimer

    from docwen_gui.release_smoke import _schedule_test_conversion_report

    warning = "缺少必需字段：成文日期、发文机关署名；识别提示：存在低置信度识别"
    source = tmp_path / "rules.docx"
    source.write_bytes(b"PK\x03\x04")
    output = tmp_path / "outputs" / "rules.md"
    output.parent.mkdir()
    output.write_text("---\n---\n", encoding="utf-8")
    report = tmp_path / "report.json"
    screenshot = tmp_path / "info-area.png"
    env = {
        "DOCWEN_GUI_TEST_CONVERSION_REPORT": str(report),
        "DOCWEN_GUI_TEST_CONVERSION_INPUT": str(source),
        "DOCWEN_GUI_TEST_CONVERSION_OUTPUT_DIR": str(output.parent),
        "DOCWEN_GUI_TEST_CONVERSION_TARGET": "md",
        "DOCWEN_GUI_TEST_CONVERSION_SURFACE": "action",
        "DOCWEN_GUI_TEST_CONVERSION_ACTION": "gongwen",
        "DOCWEN_GUI_TEST_CONVERSION_EXPECT_WARNING": warning,
        "DOCWEN_GUI_TEST_CONVERSION_SCREENSHOT": str(screenshot),
    }
    for key, value in env.items():
        monkeypatch.setenv(key, value)

    class _ConfigPort:
        def set_many(self, _values: dict[str, object]) -> bool:
            return True

    class _ViewModel:
        controller = SimpleNamespace(config_port=_ConfigPort())

        def add_files(self, paths: list[str]) -> None:
            assert paths == [str(source)]

    requests: list[str] = []

    class _ActionViewModel:
        visible = True
        action_name = ""

        def set_file_to_md_option(self, key: str, value: str) -> None:
            assert key == "optimize_for_type"
            self.action_name = value

        def request_conversion(self, target: str) -> None:
            requests.append(target)

    history_rows = [
        SimpleNamespace(message="Completed: rules.docx", message_type="warning", file_path=str(output)),
        SimpleNamespace(message=warning, message_type="warning", file_path=str(output)),
    ]
    task_summary = SimpleNamespace(
        state="success",
        tone="warning",
        completed_count=1,
        total_count=1,
        failed_count=0,
        navigate_path=str(output),
    )
    info_vm = SimpleNamespace(
        history_rows=history_rows,
        task_summary=task_summary,
        status_source="task",
        status_tone="warning",
        status_meta_text="Completed",
        status_summary_text="1/1 completed",
    )

    class _RowWidget:
        def property(self, name: str) -> str:
            assert name == "infoStatusTone"
            return "warning"

        def toolTip(self) -> str:
            return warning

        def isVisible(self) -> bool:
            return True

        def grab(self) -> _Pixmap:
            return _Pixmap()

    class _Pixmap:
        def width(self) -> int:
            return 640

        def height(self) -> int:
            return 320

        def save(self, path: str, image_format: str) -> bool:
            assert image_format == "PNG"
            Path(path).write_bytes(b"\x89PNG\r\n\x1a\nrelease-smoke")
            return True

    class _InfoArea:
        def isVisible(self) -> bool:
            return True

        def get_history_row_widget(self, index: int) -> _RowWidget | None:
            return _RowWidget() if index == 1 else None

        def grab(self) -> _Pixmap:
            return _Pixmap()

    entry = SimpleNamespace(status="completed", output_path=str(output), error_message="")
    window = SimpleNamespace(
        view_model=_ViewModel(),
        _action_area_vm=_ActionViewModel(),
        _batch_list_vm=SimpleNamespace(get_file_entry=lambda _normalized: entry),
        _info_area_vm=info_vm,
        _info_area=_InfoArea(),
        close=lambda: None,
    )
    app = SimpleNamespace(processEvents=lambda: None)
    monkeypatch.setattr(QTimer, "singleShot", lambda _delay, callback: callback())

    _schedule_test_conversion_report(cast(Any, app), cast(Any, window))

    assert requests == ["md"]
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["success"] is True
    assert payload["actionName"] == "gongwen"
    assert payload["expectedWarningMessage"] == warning
    assert payload["warningHistoryIndex"] == 1
    assert payload["warningRowTone"] == "warning"
    assert payload["warningRowTooltip"] == warning
    assert payload["warningRowVisible"] is True
    assert payload["warningRowScreenshotSaved"] is True
    assert payload["warningRowScreenshotWidth"] == 640
    assert payload["warningRowScreenshotHeight"] == 320
    assert payload["taskSummary"]["state"] == "success"
    assert payload["taskSummary"]["tone"] == "warning"
    assert payload["statusSource"] == "task"
    assert payload["statusTone"] == "warning"
    assert payload["infoAreaVisible"] is True
    assert payload["screenshotSaved"] is True
    assert payload["screenshotWidth"] == 640
    assert payload["screenshotHeight"] == 320
    assert screenshot.read_bytes().startswith(b"\x89PNG\r\n\x1a\n")
    warning_row_screenshot = screenshot.with_name("info-area_warning_row.png")
    assert warning_row_screenshot.read_bytes().startswith(b"\x89PNG\r\n\x1a\n")


def test_conversion_release_hook_stops_when_output_config_batch_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PySide6.QtCore import QTimer

    from docwen_gui.release_smoke import _schedule_test_conversion_report

    source = tmp_path / "source.md"
    source.write_text("release smoke", encoding="utf-8")
    report = tmp_path / "report.json"
    output_dir = tmp_path / "outputs"
    monkeypatch.setenv("DOCWEN_GUI_TEST_CONVERSION_REPORT", str(report))
    monkeypatch.setenv("DOCWEN_GUI_TEST_CONVERSION_INPUT", str(source))
    monkeypatch.setenv("DOCWEN_GUI_TEST_CONVERSION_OUTPUT_DIR", str(output_dir))

    batches: list[dict[str, object]] = []

    class _ConfigPort:
        def set_many(self, values: dict[str, object]) -> bool:
            batches.append(dict(values))
            return False

    class _ViewModel:
        controller = SimpleNamespace(config_port=_ConfigPort())

        def add_files(self, _paths: list[str]) -> None:
            raise AssertionError("conversion must not start after config failure")

    close_calls: list[bool] = []
    window = SimpleNamespace(view_model=_ViewModel(), close=lambda: close_calls.append(True))
    app = SimpleNamespace(processEvents=lambda: None)
    monkeypatch.setattr(QTimer, "singleShot", lambda _delay, callback: callback())

    _schedule_test_conversion_report(cast(Any, app), cast(Any, window))

    assert batches == [
        {
            "output.directory.mode": "custom",
            "output.directory.custom_path": str(output_dir),
            "output.directory.create_date_subfolder": False,
        }
    ]
    assert close_calls == [True]
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["success"] is False
    assert payload["status"] == "failed"
    assert payload["error"] == "output_config_persist_failed"
