"""Activity browsing preserves execution state and keeps one details surface."""

from dataclasses import replace

import pytest
from PySide6.QtCore import Qt
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import QApplication

from docwen_gui.dialogs.activity_records import ActivityRecordsDialog
from docwen_gui.view_models.activity_records import ActivityRecordsModel
from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
from docwen_gui.view_models.task_history import TaskHistory

pytestmark = pytest.mark.gui


@pytest.fixture
def activity(qtbot):
    history = TaskHistory()
    feedback = InfoAreaViewModel()
    model = ActivityRecordsModel(history, feedback)
    history.remember(
        {
            "request_id": "one",
            "file_path": "/source/z.md",
            "target_format": "docx",
            "options": {"spreadsheet_password": "secret-never-show"},
        }
    )
    history.record("one", "/source/z.md", "failed", "Output locked")
    history.remember({"request_id": "two", "file_path": "/source/a.csv", "target_format": "xlsx"})
    history.record("two", "/source/a.csv", "completed", output_path="/output/a.xlsx")
    first = history.get("one")
    second = history.get("two")
    assert first is not None and second is not None
    second.outcomes["/source/a.csv"] = replace(
        second.outcomes["/source/a.csv"], updated_at=first.outcomes["/source/z.md"].updated_at
    )
    history.changed.emit()
    feedback.add_message("Output locked", "danger", operation_id="one")
    dialog = ActivityRecordsDialog(model)
    qtbot.addWidget(dialog)
    dialog.show_records()
    yield history, feedback, model, dialog
    feedback.stop_all_timers()


def test_records_search_filter_sort_and_open_details(activity):
    history, _feedback, model, dialog = activity
    assert not dialog.isModal()
    assert dialog._proxy.index(0, 2).data() == "a.csv"
    assert model.rowCount() == 2  # outcome and feedback for one attempt are one record
    assert model.failed_count == 1
    assert "secret-never-show" not in " ".join(row.details for row in model.records)
    dialog.table.sortByColumn(2, Qt.SortOrder.AscendingOrder)
    assert dialog._proxy.index(0, 2).data() == "a.csv"
    dialog.show_records(failures_only=True, operation_id="one")
    assert dialog._proxy.rowCount() == 1
    assert "Output locked" in dialog.details.toPlainText()
    dialog.copy.click()
    assert QApplication.clipboard().text() == dialog.details.toPlainText()
    dialog.status_filter.setCurrentIndex(0)
    dialog.search.setText("a.xlsx")
    assert dialog._proxy.rowCount() == 1
    opened = []
    dialog.location_requested.connect(lambda path, parent: opened.append((path, parent)))
    dialog.open_output.click()
    dialog.open_source.click()
    assert opened == [("/output/a.xlsx", True), ("/source/a.csv", True)]
    dialog.search.setText("no such document")
    assert dialog._proxy.rowCount() == 0
    assert not dialog.copy.isEnabled()
    assert not dialog.open_output.isEnabled()
    assert not dialog.details.toPlainText()
    assert history.get("one").failure_details


def test_animation_does_not_reset_selected_detail_or_text_selection(activity, qtbot):
    _, feedback, _model, dialog = activity
    dialog.show_records(failures_only=True, operation_id="one")
    cursor = dialog.details.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.Start)
    cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor, 5)
    dialog.details.setTextCursor(cursor)
    selection = dialog.details.textCursor().selectedText()
    feedback.begin_task(operation_id="active", current_file="next.md", total_count=1)
    qtbot.wait(350)
    assert dialog.details.textCursor().selectedText() == selection
    assert dialog._selected().operation_id == "one"

    feedback.add_message("Unrelated notice", "info")
    assert dialog.details.textCursor().selectedText() == selection
    assert dialog._selected().operation_id == "one"


def test_clear_records_preserves_active_task_and_outputs(activity):
    history, _, model, dialog = activity
    history.remember({"request_id": "active", "file_paths": ["/one.md", "/two.md"]})
    history.record("active", "/one.md", "completed", output_path="/one.docx")
    dialog.clear.click()
    assert history.get("one") is None
    assert history.get("active").outcomes["/one.md"].output_path == "/one.docx"
    history.record("active", "/two.md", "failed", "Later failure")
    assert model.failed_count == 1
    assert any("Later failure" in row.details for row in model.records)


def test_single_success_uses_one_status_heading(activity):
    _, feedback, _, _ = activity
    feedback.set_task_summary(
        state="success",
        tone="success",
        current_file="report.md",
        output_path="/out/report.docx",
        completed_count=1,
        total_count=1,
    )
    assert "report.md" not in feedback.status_summary_text
    assert "report.docx" in feedback.output_file_text
    assert "1/1" not in feedback.status_summary_text
    assert feedback.status_meta_text not in feedback.status_summary_text


def test_batch_details_do_not_blame_another_file(activity):
    history, feedback, model, _ = activity
    history.remember({"request_id": "batch", "file_paths": ["/a.md", "/b.md"]})
    history.record("batch", "/a.md", "failed", "A is locked")
    history.record("batch", "/b.md", "completed", output_path="/b.docx")
    feedback.add_message("A is locked", "danger", operation_id="batch")
    feedback.add_message("B needs review", "warning", operation_id="batch", file_path="/b.docx")
    feedback.add_message("Batch notice", "warning", operation_id="batch")
    a = next(row for row in model.records if row.source_path == "/a.md")
    b = next(row for row in model.records if row.source_path == "/b.md")
    assert "A is locked" in a.details and "A is locked" not in b.details
    assert "B needs review" in b.details and "B needs review" not in a.details
    assert a.status == "failed" and b.status == "warning"
    assert any(row.details == "Batch notice" and not row.source_path for row in model.records)


def test_task_warnings_and_skip_reasons_survive_notice_eviction(activity):
    history, feedback, model, _ = activity
    paths = [f"/batch/{index}.md" for index in range(105)]
    history.remember({"request_id": "large", "file_paths": paths})
    for index, path in enumerate(paths):
        warning = f"Review item {index}"
        history.record("large", path, "completed", warnings=(warning,))
        feedback.add_message(warning, "warning", operation_id="large", file_path=path)
    first = next(row for row in model.records if row.source_path == paths[0])
    assert first.status == "warning" and "Review item 0" in first.details
    history.remember({"request_id": "skip", "file_path": "/skipped.md"})
    history.record("skip", "/skipped.md", "skipped", skip_reason="Unsupported in this operation")
    feedback.clear_history()
    skipped = next(row for row in model.records if row.operation_id == "skip")
    assert "Unsupported in this operation" in skipped.details
    assert next(row for row in model.records if row.source_path == paths[0]).status == "warning"


def test_file_sort_uses_visible_basename_before_directory(activity):
    history, _, _, dialog = activity
    history.remember({"request_id": "sort", "file_paths": ["/a/zebra.md", "/z/alpha.md"]})
    dialog.table.sortByColumn(2, Qt.SortOrder.AscendingOrder)
    names = [dialog._proxy.index(row, 2).data() for row in range(dialog._proxy.rowCount())]
    assert names.index("alpha.md") < names.index("zebra.md")
