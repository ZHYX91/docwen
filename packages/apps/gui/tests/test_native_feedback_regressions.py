"""Regressions discovered by the real Word and Assistant workflows."""

from pathlib import Path

import pytest
from PySide6.QtCore import QPoint
from PySide6.QtWidgets import QPushButton

from ._main_window_projection_binding_support import window as window

pytestmark = pytest.mark.gui


def test_rejected_route_replaces_previous_success_summary_without_erasing_history(window):
    from docwen_gui.diagnostics import DiagnosticSummary
    from docwen_gui.view_models.activity_records import ActivityRecordsModel

    feedback = window._info_area_vm
    feedback.set_task_summary(
        operation_id="previous",
        state="success",
        total_count=1,
        completed_count=1,
        output_path="/previous/result.md",
    )
    feedback.add_message(
        "Previous conversion completed",
        "success",
        operation_id="previous",
        diagnostic=DiagnosticSummary(status="completed", output_count=1),
    )
    model = ActivityRecordsModel(window._task_history, feedback)

    window._workflow.single(file_path="/missing/context.docx", target_format="md", action_name="", options={})

    assert feedback.output_file_text == ""
    assert not feedback.has_task_summary
    assert feedback.status_tone == "warning"
    assert feedback.history_rows[-1].message in feedback.status_summary_text
    assert feedback.history_rows[-1].diagnostic.status == "not_started"
    assert feedback.history_rows[0].message == "Previous conversion completed"
    assert any(row.status == "not_started" for row in model.records)


@pytest.mark.parametrize("mode", ["single", "batch"])
def test_ipc_new_and_repeated_open_selects_requested_input(window, tmp_path, qtbot, mode):
    first = tmp_path / "first.md"
    second = tmp_path / "second.txt"
    first.write_text("# First", encoding="utf-8")
    second.write_text("Second", encoding="utf-8")
    window._view_model.set_mode(mode)
    for source in (first, second, first, second):
        window.handle_ipc_command("open_file", str(source))
        qtbot.waitUntil(lambda: not window._view_model.inspection_busy)
        selected = window._view_model.selected_file
        assert selected is not None and Path(selected.path) == source
        assert Path(window._batch_list.get_current_file()) == source
        assert source.name in window._input_area_vm.selection_message
        context = window._view_model.ui_projection.template_context
        assert context is not None and Path(context.file_path) == source
    assert len(window._view_model.files) == (1 if mode == "single" else 2)


def test_settings_feedback_does_not_move_confirmation_buttons(qapp):
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    dialog = SettingsDialog()
    dialog.show()
    qapp.processEvents()
    buttons: list[QPushButton] = []
    for name in ("settingsOkButton", "settingsResetTabButton"):
        button = dialog.findChild(QPushButton, name)
        assert button is not None
        buttons.append(button)
    positions = [button.mapTo(dialog, QPoint()) for button in buttons]
    dialog.view_model.set_field("text", "add_numbering", not dialog.view_model.config.text.add_numbering)
    qapp.processEvents()
    assert [button.mapTo(dialog, QPoint()) for button in buttons] == positions
    for error in (False, True):
        dialog._show_status("A status message that should not move the action buttons.", error)
        qapp.processEvents()
        assert [button.mapTo(dialog, QPoint()) for button in buttons] == positions
        dialog._status_timer.timeout.emit()
        qapp.processEvents()
        assert [button.mapTo(dialog, QPoint()) for button in buttons] == positions
    dialog.view_model.cancel_changes()
    dialog.close()


def test_extension_losses_are_localized_and_summary_exposes_warning_count(qapp):
    from docwen_core.models.result import ConversionDiagnostic, ConversionResult
    from docwen_gui.execution_presenter import _result_warning_messages
    from docwen_gui.i18n import t
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel

    result = ConversionResult(
        task_id="word-off",
        success=True,
        diagnostics=[
            ConversionDiagnostic(
                level="warning",
                code="docwen.conversion.markdown_extension.typed_endnotes.flattened",
                message="Raw technical English",
                location="word/document.xml",
            )
        ],
    )
    assert _result_warning_messages(result) == [t("main_window.extension_loss_endnotes")]
    vm = InfoAreaViewModel()
    vm.set_task_summary(state="success", total_count=1, completed_count=1, warning_count=4, tone="warning")
    assert t("info_area.task_warning_count", count=4) in vm.status_summary_text
    vm.set_task_summary(state="active", total_count=1)
    assert vm.task_summary.warning_count == 0
    assert t("info_area.task_warning_count", count=4) not in vm.status_summary_text
