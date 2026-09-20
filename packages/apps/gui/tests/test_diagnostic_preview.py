"""Default copies are finite facts, and the visible preview is the copied value."""

from __future__ import annotations

import json

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from docwen_core.models.result import ConversionDiagnostic, ConversionErrorInfo, ConversionResult
from docwen_gui.diagnostics import DiagnosticSummary
from docwen_gui.dialogs.diagnostics import DiagnosticDialog
from docwen_gui.view_models.activity_records import ActivityRecordsModel
from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
from docwen_gui.view_models.task_history import TaskHistory

pytestmark = pytest.mark.gui


def test_structured_result_summary_excludes_all_free_text_and_source_identity():
    result = ConversionResult(
        task_id="secret-operation-id",
        success=False,
        error=ConversionErrorInfo(
            error_type="timeout",
            message="document secret",
            traceback_text="/private/source.py",
            diagnostic_code="TOKEN_LOOKING_LIKE_A_CODE",
            recoverable=True,
        ),
        diagnostics=[ConversionDiagnostic(level="warning", message="body secret", location="/private/input")],
    )
    snapshot = DiagnosticSummary.from_result(result, output_count=2)
    assert result.error is not None
    result.error.error_type = "secret replacement"
    result.diagnostics.clear()
    assert json.loads(snapshot.to_text()) == {
        "schema": "docwen.gui-diagnostic.v1",
        "status": "failed",
        "error_category": "timeout",
        "output_count": 2,
        "warning_count": 1,
        "reported_recoverable": True,
    }


def test_unknown_labels_and_exception_names_cannot_smuggle_data():
    class SECRET_EXCEPTION(RuntimeError):
        pass

    assert (
        json.loads(DiagnosticSummary.from_exception(SECRET_EXCEPTION("secret")).to_text())["exception_type"]
        == "unknown"
    )
    snapshot = DiagnosticSummary(
        status="private", error_category="private", exception_type="private", output_count=-5, warning_count=True
    )
    assert json.loads(snapshot.to_text()) == {
        "schema": "docwen.gui-diagnostic.v1",
        "status": "unknown",
        "output_count": 0,
        "warning_count": 0,
        "error_category": "unknown",
        "exception_type": "unknown",
    }


def test_diagnostic_copy_keeps_dialog_open_and_local_details_separate(qtbot):
    dialog = DiagnosticDialog(
        "Error", "Task explanation", details="/private/input secret", diagnostic=DiagnosticSummary(status="failed")
    )
    qtbot.addWidget(dialog)
    QApplication.clipboard().setText("clipboard untouched")
    dialog.show()
    assert QApplication.clipboard().text() == "clipboard untouched"
    assert dialog.view.currentIndex() == 1
    qtbot.mouseClick(dialog.copy, Qt.MouseButton.LeftButton)
    assert dialog.isVisible()
    assert QApplication.clipboard().text() == dialog.view.preview.toPlainText()
    assert "secret" not in QApplication.clipboard().text()
    assert dialog.view.details.toPlainText() == "/private/input secret"


@pytest.mark.parametrize("batch", [False, True])
def test_presenter_preserves_error_facts_for_activity_without_copying_context(main_window, tmp_path, batch):
    source = tmp_path / "private.md"
    source.write_text("secret", encoding="utf-8")
    context = {
        "request_id": "private-task",
        "file_path": str(source),
        "file_paths": [str(source)],
        "batch": batch,
        "total_count": 1,
        "options": {"password": "secret"},
    }
    result = ConversionResult(
        task_id="private-task",
        success=False,
        error=ConversionErrorInfo(error_type="dependency_missing", message="private command", recoverable=False),
    )
    main_window._results.finished([result] if batch else result, context)
    record = main_window._task_history.get("private-task")
    assert record is not None
    outcome = record.outcomes[source.as_posix()]
    assert outcome.diagnostic is not None
    payload = json.loads(outcome.diagnostic.to_text())
    assert payload["error_category"] == "dependency_missing"
    assert payload["reported_recoverable"] is False
    assert "private" not in outcome.diagnostic.to_text()
    assert outcome.error_message == "private command"


def test_notice_and_legacy_history_copies_have_no_unstructured_data(qapp):
    history = TaskHistory()
    feedback = InfoAreaViewModel()
    model = ActivityRecordsModel(history, feedback)
    feedback.add_message("raw secret /private/path", "warning", file_path="/private/path")
    history.remember({"request_id": "secret", "file_path": "/private/input", "options": {"key": "secret"}})
    history.record("secret", "/private/input", "failed", "raw secret")
    assert len(model.records) == 2
    for record in model.records:
        assert record.diagnostic is not None
        text = record.diagnostic.to_text()
        assert "private" not in text and "secret" not in text
    feedback.stop_all_timers()


def test_batch_diagnostic_action_opens_exact_record_without_copying_raw_error(main_window, tmp_path):
    from docwen_gui.dialogs.activity_records import ActivityRecordsDialog

    source = tmp_path / "source.md"
    source.write_text("private content", encoding="utf-8")
    paths, rejected = main_window._batch_list_vm.add_files([str(source)])
    assert not rejected
    path = paths[0]
    main_window._task_history.remember({"request_id": "one", "file_path": path})
    main_window._results.file_status(path, "failed", operation_id="one", error_message="private content")
    QApplication.clipboard().setText("untouched")
    main_window._handle_batch_entry_action("show_diagnostics", path)
    dialog = main_window.findChild(ActivityRecordsDialog)
    assert dialog is not None
    selected = dialog._selected()
    assert selected is not None and selected.operation_id == "one"
    assert dialog.diagnostic_view.currentIndex() == 1
    assert QApplication.clipboard().text() == "untouched"
    dialog.copy.click()
    assert QApplication.clipboard().text() == dialog.diagnostic_view.preview.toPlainText()
    assert "private" not in QApplication.clipboard().text()
