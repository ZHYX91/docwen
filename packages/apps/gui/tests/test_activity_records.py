"""Activity browsing preserves execution state and keeps one details surface."""

import json
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


def test_failure_details_keep_diagnostics_without_self_navigation(qapp):
    from docwen_gui.i18n import t

    history = TaskHistory()
    feedback = InfoAreaViewModel()
    model = ActivityRecordsModel(history, feedback)
    history.remember({"request_id": "failure", "file_path": "/input.md", "target_format": "docx"})
    history.record("failure", "/input.md", "failed", "Technical note incompatibility")
    notice = t("main_window.conversion_failed") + " [MD2DOCX-NOTE-PART-INVALID]"
    feedback.add_message(
        notice,
        "danger",
        operation_id="failure",
        details_message=t("main_window.conversion_failed_detail") + " [MD2DOCX-NOTE-PART-INVALID]",
    )
    assert feedback.history_rows[0].message == notice
    details = model.records[0].details
    assert t("main_window.conversion_failed") not in details
    assert "Technical note incompatibility" in details
    assert "MD2DOCX-NOTE-PART-INVALID" in details
    feedback.stop_all_timers()


@pytest.mark.parametrize("locale", ["zh_CN", "en_US", "ar_SA"])
def test_note_failures_have_localized_specific_diagnoses(qapp, locale):
    from docwen_core.models.result import ConversionErrorInfo
    from docwen_gui.execution_presenter import _localized_failure_message
    from docwen_gui.i18n import get_locale, set_locale, t

    previous = get_locale()
    set_locale(locale)
    try:
        for code, key in (
            ("MD2DOCX-NOTE-SYNTAX-INVALID", "main_window.note_syntax_invalid"),
            ("MD2DOCX-NOTE-PART-INVALID", "main_window.note_template_incompatible"),
        ):
            error = ConversionErrorInfo(error_type="invalid_input", message="Technical details", diagnostic_code=code)
            assert _localized_failure_message(error) == t(key)
        error = ConversionErrorInfo(
            error_type="conversion_failed", message="disk error", diagnostic_code="MD2DOCX-NOTE-WRITEBACK-ERROR"
        )
        assert _localized_failure_message(error) != t("main_window.note_template_incompatible")
    finally:
        set_locale(previous)


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
    copied = QApplication.clipboard().text()
    assert copied == dialog.diagnostic_view.preview.toPlainText()
    assert json.loads(copied)["status"] == "failed"
    assert "Output locked" not in copied and "/source" not in copied and "secret" not in copied
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


def test_merge_locations_and_request_summary_survive_notice_eviction_without_becoming_copyable(activity):
    history, feedback, model, _dialog = activity
    history.remember(
        {
            "request_id": "merge",
            "file_paths": ["/base.xlsx", "/collect.xlsx"],
            "aggregate": True,
            "target_format": "xlsx",
            "request_summary": "Template: Report · Output: /exports",
        }
    )
    history.record("merge", "/base.xlsx", "completed")
    history.record("merge", "/collect.xlsx", "completed", local_details=("added=1; source Sheet!row 2",))
    feedback.clear_history()
    collection = next(record for record in model.records if record.source_path == "/collect.xlsx")
    assert "source Sheet!row 2" in collection.details
    assert "Template: Report" in collection.details
    assert collection.diagnostic is not None
    copied = collection.diagnostic.to_text()
    assert all(value not in copied for value in ("Sheet", "Report", "/exports"))


def test_presenter_attaches_merge_decisions_to_the_matching_collection(main_window):
    from docwen_core.models.result import ConversionDiagnostic, ConversionMetrics, ConversionResult

    paths = ["/base.xlsx", "/first.xlsx", "/second.xlsx"]
    context = {
        "request_id": "merge",
        "file_path": paths[0],
        "file_paths": paths,
        "aggregate": True,
        "total_count": 3,
        "target_format": "xlsx",
        "display_name": "base.xlsx",
        "request_summary": "Output: /exports",
    }
    result = ConversionResult(
        task_id="merge",
        success=True,
        diagnostics=[
            ConversionDiagnostic(level="info", code="MERGE-SUMMARY", message="first: added row 2"),
            ConversionDiagnostic(level="info", code="MERGE-SUMMARY", message="second: replaced row 4"),
        ],
        metrics=ConversionMetrics(
            extra={
                "merge_summaries": [
                    {
                        "source_name": "first.xlsx",
                        "sheet": "Sheet",
                        "mode": "row",
                        "counts": {"added": 1, "replaced": 0, "skipped": 0, "combined": 0},
                    },
                    {
                        "source_name": "second.xlsx",
                        "sheet": "Sheet",
                        "mode": "row",
                        "counts": {"added": 0, "replaced": 1, "skipped": 0, "combined": 0},
                    },
                ]
            }
        ),
    )
    main_window._results.finished(result, context)
    main_window._info_area_vm.clear_history()
    record = main_window._task_history.get("merge")
    assert record is not None
    assert record.outcomes[paths[0]].local_details == ()
    assert record.outcomes[paths[1]].local_details == ("first: added row 2",)
    assert record.outcomes[paths[2]].local_details == ("second: replaced row 4",)
    assert "first.xlsx" in main_window._info_area_vm.execution_summary
    assert "second.xlsx" in main_window._info_area_vm.execution_summary


@pytest.mark.parametrize(
    "malformed",
    [None, {}, "summary", {"counts": {}}, {"added": -1}, {"added": "1"}, {"added": True}, {"added": 2**100}],
)
def test_malformed_optional_merge_metrics_do_not_block_success(main_window, malformed):
    from docwen_core.models.result import ConversionMetrics, ConversionResult

    summary = {
        "source_name": "collect.xlsx",
        "sheet": "Sheet",
        "mode": "row",
        "counts": {"added": 1, "replaced": 0, "skipped": 0, "combined": 0},
    }
    if isinstance(malformed, dict) and "added" in malformed:
        summary["counts"].update(malformed)
        malformed = summary
    paths = ["/base.xlsx", "/collect.xlsx"]
    context = {
        "request_id": "malformed-metrics",
        "file_path": paths[0],
        "file_paths": paths,
        "aggregate": True,
        "total_count": 2,
        "target_format": "xlsx",
        "display_name": "base.xlsx",
    }
    result = ConversionResult(
        task_id="malformed-metrics",
        success=True,
        metrics=ConversionMetrics(extra={"merge_summaries": [malformed]}),
    )
    main_window._results.finished(result, context)
    record = main_window._task_history.get("malformed-metrics")
    assert record is not None
    assert all(record.outcomes[path].status == "completed" for path in paths)
    assert main_window._info_area_vm.task_summary.state == "success"
    assert main_window._info_area_vm.execution_summary == ""


def test_merge_metrics_render_only_the_actual_collection_count(main_window):
    from docwen_core.models.result import ConversionMetrics, ConversionResult

    result = ConversionResult(
        task_id="bounded",
        success=True,
        metrics=ConversionMetrics(
            extra={
                "merge_summaries": [
                    {
                        "source_name": source,
                        "sheet": "Sheet",
                        "mode": "row",
                        "counts": {"added": 1, "replaced": 0, "skipped": 0, "combined": 0},
                    }
                    for source in ("collect.xlsx", "unexpected.xlsx")
                ],
            }
        ),
    )
    main_window._results._show_merge_summary(result, {"aggregate": True, "file_paths": ["/base.xlsx", "/collect.xlsx"]})
    assert "collect.xlsx" in main_window._info_area_vm.execution_summary
    assert "unexpected.xlsx" not in main_window._info_area_vm.execution_summary


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


def test_preflight_notices_keep_file_operation_reason_and_safe_copy(qtbot):
    from docwen_gui.diagnostics import DiagnosticSummary

    history, feedback = TaskHistory(), InfoAreaViewModel()
    model = ActivityRecordsModel(history, feedback)
    facts = DiagnosticSummary(
        status="not_started",
        phase="preflight",
        diagnostic_code="ROUTE-NOT-AVAILABLE",
        source_format="wps",
        target_format="md",
    )
    for _ in range(2):
        feedback.add_message("No route", "warning", file_path="/private/one.wps", operation="Gongwen", diagnostic=facts)
    feedback.add_message("No route", "warning", file_path="/private/two.wps", operation="Gongwen", diagnostic=facts)
    assert len(model.records) == 2
    assert all(row.status == "not_started" for row in model.records)
    assert model.failed_count == 0
    assert feedback.history_rows[0].repeat_count == 2
    copied = facts.to_text()
    assert json.loads(copied)["phase"] == "preflight"
    assert "/private/" not in copied
    unsafe = DiagnosticSummary(
        status="not_started",
        phase="preflight",
        diagnostic_code="secret-user-value",
        source_format="private-path",
        target_format="private-target",
    )
    assert "private" not in unsafe.to_text() and "secret" not in unsafe.to_text()
    feedback.stop_all_timers()
