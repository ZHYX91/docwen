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
    from PySide6.QtCore import QObject, QTimer, Signal

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
        window_holder["window"]._execution.result_ready.emit(result, {"surface": request_surface})

    def original_execution_finished(observed_result: object, context: dict[str, object]) -> None:
        assert observed_result is result
        handler_calls.append(context)
        entry.status = "completed"

    class _Execution(QObject):
        result_ready = Signal(object, dict)

    execution = _Execution()
    execution.result_ready.connect(original_execution_finished)
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
        _execution=execution,
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


@pytest.mark.parametrize("warning_owner", ["current", "other", "missing"])
def test_conversion_release_hook_captures_rendered_successful_warning(
    qapp, qtbot, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, warning_owner: str
) -> None:
    from PySide6.QtCore import QTimer

    from docwen_gui.dialogs.activity_records import ActivityRecordsDialog
    from docwen_gui.release_smoke import _schedule_test_conversion_report
    from docwen_gui.view_models.activity_records import ActivityRecordsModel
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
    from docwen_gui.view_models.task_history import TaskHistory
    from docwen_gui.widgets.info_area import InfoArea

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
    history = TaskHistory()
    info_vm = InfoAreaViewModel()
    for operation in ("current", "other"):
        history.remember(
            {
                "request_id": operation,
                "file_path": str(source),
                "target_format": "md",
                "action_name": "gongwen",
                "request_summary": "Gongwen request summary",
            }
        )
        history.record(
            operation,
            str(source),
            "completed",
            output_path=str(output),
            warnings=(warning,) if warning_owner == operation else (),
            local_details=("Local extraction details",),
        )
    if warning_owner != "missing":
        info_vm.add_message(warning, "warning", file_path=str(output), operation_id=warning_owner)
    # The unrelated operation may have the same text; it must never satisfy
    # the current operation's smoke evidence or select its details row.
    info_vm.set_task_summary(
        operation_id="current",
        state="success",
        tone="warning",
        completed_count=1,
        total_count=1,
        output_path=str(output),
        navigate_file_path=str(output),
        warning_count=1,
    )
    info_vm.clear_all_transients()
    model = ActivityRecordsModel(history, info_vm)
    dialog = ActivityRecordsDialog(model)
    info_area = InfoArea(info_vm)
    qtbot.addWidget(dialog)
    qtbot.addWidget(info_area)
    info_area.resize(700, 250)
    info_area.show()
    qapp.processEvents()
    current = next(row for row in model.records if row.operation_id == "current")
    assert current.source_path == source.as_posix()
    assert current.output_path == str(output)
    assert "Gongwen request summary" in current.details
    assert "Local extraction details" in current.details
    assert str(source) in current.details and str(output) in current.details
    assert current.details != warning
    requests = []

    class _ViewModel:
        controller = SimpleNamespace(config_port=SimpleNamespace(set_many=lambda _values: True))

        def add_files(self, paths):
            assert paths == [str(source)]

    class _ActionViewModel:
        visible = True
        action_name = ""

        def set_file_to_md_option(self, key, value):
            assert key == "optimize_for_type"
            self.action_name = value

        def request_conversion(self, target):
            requests.append(target)

    def show_records(**kwargs):
        dialog.show_records(**kwargs)
        return dialog

    entry = SimpleNamespace(status="completed", output_path=str(output), error_message="", operation_id="current")
    window = SimpleNamespace(
        view_model=_ViewModel(),
        _action_area_vm=_ActionViewModel(),
        _batch_list_vm=SimpleNamespace(get_file_entry=lambda _normalized: entry),
        _info_area_vm=info_vm,
        _info_area=info_area,
        _show_activity_records=show_records,
        close=lambda: None,
    )
    monkeypatch.setattr(QTimer, "singleShot", lambda _delay, callback: callback())
    try:
        _schedule_test_conversion_report(qapp, cast(Any, window))
        assert requests == ["md"]
        payload = json.loads(report.read_text(encoding="utf-8"))
        assert payload["success"] is (warning_owner == "current")
        if warning_owner != "current":
            assert payload["error"] == "successful_warning_ui_mismatch"
            assert payload["warningHistoryIndex"] == -1
            return
        assert payload["warningRecord"] == {
            "operationId": "current",
            "sourcePath": source.as_posix(),
            "outputPath": str(output),
            "warnings": [warning],
            "details": current.details,
        }
        assert payload["warningDetailsText"] == current.details
        assert payload["warningRecordStatus"] == "warning"
        assert payload["warningDetailsVisible"] is True
        assert payload["warningDetailsScreenshotSaved"] is True
        assert payload["warningDetailsScreenshotWidth"] > 0
        assert payload["warningDetailsScreenshotHeight"] > 0
        assert payload["taskSummary"]["state"] == "success"
        assert payload["statusSource"] == "task"
        assert payload["statusTone"] == "warning"
        assert screenshot.read_bytes().startswith(b"\x89PNG\r\n\x1a\n")
    finally:
        info_vm.stop_all_timers()


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
