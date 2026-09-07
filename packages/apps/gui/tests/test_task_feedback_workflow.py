"""User-event regressions for task identity, result actions and compact feedback."""

from __future__ import annotations

import threading
from types import SimpleNamespace

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel

from docwen_core.events.task_events import TASK_PROGRESS
from docwen_gui.main_window import _normalize_path
from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
from docwen_gui.widgets.info_area import InfoArea

pytestmark = pytest.mark.gui


def test_next_worker_replaces_old_result_and_ignores_late_old_telemetry(main_window, qtbot):
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    release = threading.Event()

    class Controller:
        has_runtime = True

        def prepare_execution_cancellation(self, request, *, batch=False):
            return object()

        def execute_single(self, request):
            release.wait(3)
            return ConversionResult(
                task_id=request.request_id,
                success=False,
                error=ConversionErrorInfo(error_type="cancelled", message="Cancelled"),
            )

        def release_execution_cancellation(self, *args):
            pass

    window = main_window
    vm = window._info_area_vm
    vm.set_task_summary(
        operation_id="old",
        state="failed",
        current_file="old.docx",
        guide_actions=vm.compute_guide_actions("failed", failed_details_path="old.docx", retry_available=True),
    )
    context = {
        "request_id": "new",
        "file_path": "new.docx",
        "file_paths": ["new.docx"],
        "display_name": "new.docx",
        "total_count": 1,
    }
    try:
        assert window._launch_execution_thread(
            controller=Controller(),
            request=SimpleNamespace(request_id="new"),
            context=context,
            project_reserved_execution=lambda: window.view_model.begin_execution_telemetry("new", ("new",)),
        )
        assert vm.task_summary.operation_id == "new"
        assert vm.task_summary.state == "active"
        assert not vm.guide_actions
        assert not vm.status_action_target
        window.view_model.on_task_event(TASK_PROGRESS, {"task_id": "new", "message": "Writing", "percent": 25})
        assert vm.task_summary.percent == 25
        window.view_model.on_task_event(TASK_PROGRESS, {"task_id": "old", "message": "Obsolete", "percent": 99})
        assert vm.task_summary.percent == 25
        vm.set_transient_message("font_size", "Font: Large", ttl_ms=1000)
        assert vm.notification_text == "Font: Large"
        assert "new.docx" in vm.status_summary_text
        vm.mark_cancelling("new")
        window.view_model.on_task_event(TASK_PROGRESS, {"task_id": "new", "message": "Late progress", "percent": 80})
        assert vm.task_summary.state == "cancelling"
        assert vm.task_summary.percent == 25
    finally:
        release.set()
        qtbot.waitUntil(lambda: not window._active_threads, timeout=5000)


@pytest.mark.parametrize("kind", ["single", "batch", "aggregate"])
def test_retry_preserves_scope_order_and_options(main_window, tmp_path, monkeypatch, kind):
    paths = []
    for name in ("b", "a", "other"):
        path = tmp_path / f"{name}.md"
        path.write_text("# source", encoding="utf-8")
        paths.append(_normalize_path(str(path)))
    main_window._batch_list_vm.add_files(paths)
    for path in paths:
        main_window._batch_list_vm.set_file_status(
            path, "failed", operation_id="op" if path != paths[2] else "unrelated"
        )
    context = {
        "request_id": "op",
        "file_path": paths[0],
        "file_paths": paths[:2],
        "target_format": "docx",
        "action_name": "convert",
        "options": {"template_name": "chosen"},
    }
    if kind != "single":
        context[kind] = True
    main_window._task_history.remember(context)
    for path in paths[:2]:
        main_window._task_history.record("op", path, "failed")
    main_window._info_area_vm.set_task_summary(operation_id="op", state="failed")
    calls = []
    for method in ("_start_execution", "_start_batch_execution", "_start_aggregate_execution"):
        monkeypatch.setattr(main_window, method, lambda _method=method, **kwargs: calls.append((_method, kwargs)))
    main_window._retry_failed_request()
    expected = {
        "single": "_start_execution",
        "batch": "_start_batch_execution",
        "aggregate": "_start_aggregate_execution",
    }[kind]
    assert len(calls) == 1
    assert calls[0][0] == expected
    assert calls[0][1]["options"] == context["options"]
    if kind == "single":
        assert calls[0][1]["file_path"] == paths[0]
    else:
        assert calls[0][1]["file_paths"] == paths[:2]
    assert len(main_window._batch_list_vm.get_failed_files()) == 3  # no state reset before admission


def test_viewing_details_does_not_republish_failures(main_window, tmp_path):
    from docwen_gui.dialogs.task_details import TaskDetailsDialog

    path = tmp_path / "report.md"
    path.write_text("# source", encoding="utf-8")
    main_window._batch_list_vm.add_files([str(path)])
    main_window._batch_list_vm.set_file_status(str(path), "failed", error_message="Output is locked", operation_id="op")
    main_window._task_history.remember({"request_id": "op", "file_path": str(path)})
    main_window._task_history.record("op", str(path), "failed", "Output is locked")
    main_window._info_area_vm.set_task_summary(operation_id="op", state="failed")
    before = main_window._info_area_vm.history_rows
    main_window._handle_task_guide_action("view_failed_details", str(path))
    main_window._handle_task_guide_action("view_failed_details", str(path))
    dialogs = main_window.findChildren(TaskDetailsDialog)
    assert len(dialogs) == 1
    assert "Output is locked" in dialogs[0].details.toPlainText()
    assert main_window._info_area_vm.history_rows == before
    dialogs[0].close()


@pytest.mark.parametrize("accepted", [True, False])
def test_retry_requests_a_fresh_password_without_storing_it(main_window, tmp_path, monkeypatch, accepted):
    source = tmp_path / "encrypted.xlsx"
    source.write_bytes(b"fixture")
    path = _normalize_path(str(source))
    main_window._batch_list_vm.add_files([path])
    main_window._batch_list_vm.set_file_status(path, "failed", operation_id="encrypted")
    main_window._info_area_vm.set_task_summary(operation_id="encrypted", state="failed")
    options = {"spreadsheet_password": "<redacted>", "allow_protection_loss": True}
    main_window._task_history.remember(
        {
            "request_id": "encrypted",
            "file_path": path,
            "target_format": "ods",
            "action_name": "convert",
            "options": options,
        }
    )
    main_window._task_history.record("encrypted", path, "failed")
    monkeypatch.setattr("docwen_gui.main_window.QInputDialog.getText", lambda *args: ("fresh-password", accepted))
    calls = []
    monkeypatch.setattr(main_window, "_start_execution", lambda **kwargs: calls.append(kwargs))
    main_window._retry_failed_request()
    assert len(calls) == int(accepted)
    if accepted:
        assert calls[0]["options"]["spreadsheet_password"] == "fresh-password"
    assert options["spreadsheet_password"] == "<redacted>"


def test_animation_keeps_history_scroll_selection_and_row_identity(qtbot):
    vm = InfoAreaViewModel()
    widget = InfoArea(vm)
    qtbot.addWidget(widget)
    widget.resize(460, 440)
    widget.show()
    for index in range(35):
        vm.add_message(f"Failure {index}: review this output before retrying.", "danger")
    widget._history_toggle.setChecked(True)
    scrollbar = widget._scroll.verticalScrollBar()
    qtbot.waitUntil(lambda: scrollbar.maximum() > 0)
    qtbot.wait(150)
    scrollbar.setValue(0)
    row = widget.get_history_row_widget(0)
    assert row is not None
    label = row.findChild(QLabel, "infoHistoryText")
    assert label is not None
    label.setSelection(0, 7)
    selected = label.selectedText()
    vm.begin_task(operation_id="active", current_file="next.md", total_count=1)
    qtbot.wait(400)
    assert widget.get_history_row_widget(0) is row
    assert label.selectedText() == selected
    assert scrollbar.value() == 0
    vm.stop_all_timers()


def test_history_follow_does_not_scroll_past_the_last_message(qtbot):
    vm = InfoAreaViewModel()
    widget = InfoArea(vm)
    qtbot.addWidget(widget)
    widget.resize(240, 420)
    widget.show()
    for index in range(8):
        vm.add_message(f"{index}: A wrapped message with enough detail to span several lines.", "info")
    widget._history_toggle.setChecked(True)
    qtbot.wait(100)
    last = widget._history_row_widgets[-1]
    widget._msg_container.setMinimumHeight(last.geometry().bottom() + 400)
    qtbot.wait(30)
    widget._scroll_to_bottom()
    viewport = widget._scroll.viewport()
    assert 0 < last.mapTo(viewport, last.rect().bottomRight()).y() <= viewport.height()
    assert widget._scroll.verticalScrollBar().value() < widget._scroll.verticalScrollBar().maximum()
    vm.stop_all_timers()


@pytest.mark.parametrize("activation", ["mouse", "return", "space"])
def test_summary_navigation_uses_real_user_events(qtbot, activation):
    vm = InfoAreaViewModel()
    widget = InfoArea(vm)
    qtbot.addWidget(widget)
    vm.set_task_summary(state="success", current_file="report.docx", navigate_file_path="/output/report.docx")
    widget.resize(460, 250)
    widget.show()
    button = widget._status_summary_label
    emitted = []
    vm.history_navigation_requested.connect(emitted.append)
    if activation == "mouse":
        qtbot.mouseClick(button, Qt.MouseButton.LeftButton)
    else:
        button.setFocus()
        qtbot.keyClick(button, Qt.Key.Key_Return if activation == "return" else Qt.Key.Key_Space)
    assert emitted == ["/output/report.docx"]
    vm.stop_all_timers()


def test_empty_failure_actions_do_not_offer_dead_details():
    assert InfoAreaViewModel.compute_guide_actions("failed") == []


def test_batch_progress_spans_children_without_resetting_to_zero(main_window):
    from docwen_core.events.task_events import TASK_COMPLETED, TASK_STARTED

    main_window._feedback_context = {"request_id": "batch", "batch": True, "file_paths": ["a.md", "b.md"]}
    vm = main_window._info_area_vm
    vm.begin_task(operation_id="batch", current_file="Two files", total_count=2)
    main_window.view_model.begin_execution_telemetry("batch", ("batch-0", "batch-1"))
    for event, task, percent, expected in (
        (TASK_PROGRESS, "batch-0", 50, 25),
        (TASK_COMPLETED, "batch-0", None, 50),
        (TASK_STARTED, "batch-1", None, 50),
        (TASK_PROGRESS, "batch-1", 50, 75),
    ):
        main_window.view_model.on_task_event(event, {"task_id": task, "percent": percent})
        assert vm.task_summary.percent == expected
    assert vm.task_summary.current_file == "b.md"
    assert vm.task_summary.completed_count == 1


def test_all_skipped_batch_does_not_claim_failure(main_window, tmp_path):
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    source = tmp_path / "skip.md"
    source.write_text("# source", encoding="utf-8")
    main_window._batch_list_vm.add_files([str(source)])
    main_window._on_execution_finished(
        [
            ConversionResult(
                task_id="skip-0", success=False, error=ConversionErrorInfo(error_type="skipped", message="Skipped")
            )
        ],
        {"request_id": "skip", "file_paths": [str(source)], "total_count": 1, "batch": True},
    )
    summary = main_window._info_area_vm.task_summary
    assert summary.state == "skipped"
    assert summary.failed_count == 0
    assert summary.skipped_count == 1
    assert not main_window._info_area_vm.guide_actions


def test_input_selection_updates_merge_reference_and_availability(main_window_with_controller, tmp_path):
    main_window = main_window_with_controller
    from openpyxl import Workbook

    paths = [tmp_path / name for name in ("first.xlsx", "reference.xlsx")]
    for path in paths:
        Workbook().save(path)
    main_window._input_area_vm.add_files([str(paths[0])])
    assert main_window._conversion_panel._merge_tables_button is None
    assert main_window._conversion_panel._extra_group.isHidden()
    main_window._input_area_vm.set_mode("batch")
    button = main_window._conversion_panel._merge_tables_button
    assert button is not None and not button.isEnabled()
    assert button.toolTip()
    main_window._input_area_vm.add_files([str(paths[1])])
    main_window._on_selected_file_changed(str(paths[1]))
    assert main_window._conversion_panel_vm.reference_table_name == "reference.xlsx"
    assert "reference.xlsx" in main_window._input_area_vm.selection_message
    button = main_window._conversion_panel._merge_tables_button
    assert button is not None and button.isEnabled()
    assert not button.toolTip()


def test_batch_category_selection_updates_input_summary_and_conversion_target(
    main_window_with_controller, tmp_path, qtbot
):
    from PIL import Image

    window = main_window_with_controller
    text = tmp_path / "source.md"
    text.write_text("# Source", encoding="utf-8")
    image = tmp_path / "current.png"
    Image.new("RGB", (24, 24), "white").save(image)
    window._input_area_vm.set_mode("batch")
    window._input_area_vm.add_files([str(text), str(image)])
    window._batch_list._activate_tab("markdown")
    assert "source.md" in window._input_area_vm.selection_message
    window._batch_list._activate_tab("image")
    qtbot.wait(30)
    assert "current.png" in window._input_area_vm.selection_message
    assert "source.md" not in window._input_area_vm.selection_message
    assert window.view_model.selected_file.path == str(image)
    assert window._conversion_panel_vm.current_file_path == str(image)


def test_form_row_label_and_control_do_not_overlap_when_resized(qtbot):
    from PySide6.QtWidgets import QComboBox

    from docwen_gui.widgets.panel_card import FormRow

    control = QComboBox()
    control.addItem("PDF")
    control.setMinimumWidth(120)
    row = FormRow("Target format", control)
    qtbot.addWidget(row)
    row.show()
    for width in (300, 150, 300):
        row.resize(width, 90)
        qtbot.wait(30)
        assert not row.label.geometry().intersects(control.geometry())
        assert row.rect().contains(row.label.geometry())
        assert row.rect().contains(control.geometry())


def test_required_password_is_labelled_as_required(qtbot, monkeypatch):
    from docwen_gui.i18n import t
    from docwen_gui.spreadsheet_protection import SpreadsheetProtectionInfo
    from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
    from docwen_gui.widgets.conversion_panel import ConversionPanel

    monkeypatch.setattr(
        "docwen_gui.view_models.spreadsheet_analysis._Inspector.inspect",
        lambda self, path: SpreadsheetProtectionInfo(
            path=path, status="protected", protected_parts=("encrypted-package",)
        ),
    )
    from tests.support.gui_vm_fakes import FakeMainWindowViewModel

    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(vm)
    qtbot.addWidget(widget)
    vm.set_file_info("spreadsheet", "xlsx", file_path="encrypted.xlsx")
    qtbot.waitUntil(lambda: not vm.spreadsheet_analysis_pending)
    edit = widget._spreadsheet_password_edit
    assert edit is not None
    assert edit.placeholderText() == t("conversion_panel.spreadsheet.protection_password_required")
    assert edit.accessibleName() == edit.placeholderText()


def test_extremely_narrow_panel_keeps_overflow_reachable(qtbot):
    from PySide6.QtWidgets import QScrollArea
    from tests.support.gui_vm_fakes import FakeMainWindowViewModel

    from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
    from docwen_gui.widgets.conversion_panel import ConversionPanel

    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(vm)
    qtbot.addWidget(widget)
    vm.set_file_info("spreadsheet", "xlsx", file_path="report.xlsx")
    widget.resize(130, 400)
    widget.show()
    qtbot.wait(100)
    scroll = widget.findChild(QScrollArea, "conversionPanelScrollArea")
    assert scroll is not None
    bar = scroll.horizontalScrollBar()
    assert bar.maximum() == 0 or bar.isVisible()
    bar.setValue(bar.maximum())
    content = scroll.widget()
    assert content is not None
    assert content.mapTo(scroll.viewport(), content.rect().bottomRight()).x() <= scroll.viewport().width()


def test_rebuilding_protection_controls_cancels_deleted_label_callbacks(qtbot):
    from PySide6.QtCore import QCoreApplication, QEvent
    from tests.support.gui_vm_fakes import FakeMainWindowViewModel

    from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
    from docwen_gui.widgets.conversion_panel import ConversionPanel

    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(vm)
    qtbot.addWidget(widget)
    for _ in range(3):
        row, _checkbox = widget._make_wrapping_checkbox("Allow protection loss")
        row.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    qtbot.wait(30)


def test_settings_short_window_keeps_navigation_and_confirmation_reachable(qtbot):
    from PySide6.QtWidgets import QPushButton

    from docwen_gui.widgets.settings.dialog import SettingsDialog

    dialog = SettingsDialog()
    qtbot.addWidget(dialog)
    dialog.show()
    dialog.resize(510, 560)
    qtbot.wait(100)
    assert dialog.height() <= 560
    assert dialog._compact_navigation.isVisible()
    dialog._compact_navigation.setCurrentIndex(5)
    assert dialog._tab_widget.currentIndex() == 5
    for name in ("settingsOkButton", "settingsCancelButton", "settingsApplyButton"):
        button = dialog.findChild(QPushButton, name)
        assert button is not None
        assert button.isVisible()
        assert dialog.rect().contains(button.mapTo(dialog, button.rect().bottomRight()))
    dialog.close()
