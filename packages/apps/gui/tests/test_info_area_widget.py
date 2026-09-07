"""Smoke tests for InfoArea widget.

Tests widget construction, history rendering, transient display,
task summary, guide buttons, and location button interaction.
Requires a QApplication instance.
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest
from PySide6.QtWidgets import (
    QApplication,
)

from docwen_gui.view_models.info_area_vm import (
    InfoAreaViewModel,
)
from docwen_gui.widgets.info_area import InfoArea

pytestmark = pytest.mark.gui


def test_result_filename_is_elided_with_full_output_tooltip(qtbot, qapp) -> None:
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
    from docwen_gui.widgets.info_area import InfoArea

    vm = InfoAreaViewModel()
    widget = InfoArea(view_model=vm)
    qtbot.addWidget(widget)
    widget.setFixedWidth(340)
    output = "/outputs/" + "很长的实际输出名称" * 8 + "_20260907_180000_fromMd.docx"
    vm.set_task_summary(current_file="input.md", output_path=output, state="success", completed_count=1, total_count=1)
    widget.show()
    qapp.processEvents()
    label = widget._output_row.name_label
    assert label.isVisible()
    assert label.toolTip() == output
    assert label.full_text.endswith("_20260907_180000_fromMd.docx")
    assert label.text() != label.full_text
    assert not widget._status_summary_label.isVisible()
    vm.begin_task(operation_id="next", current_file="next.md", total_count=1)
    assert not label.isVisible()


@pytest.fixture
def vm() -> Iterator[InfoAreaViewModel]:
    v = InfoAreaViewModel()
    yield v
    v.stop_all_timers()


@pytest.fixture
def widget(qapp: QApplication, vm: InfoAreaViewModel) -> Iterator[InfoArea]:
    w = InfoArea(view_model=vm)
    yield w
    w.deleteLater()


# ── Construction ──────────────────────────────────────────────────────────


class TestConstruction:
    def test_widget_created(self, widget: InfoArea) -> None:
        assert widget is not None

    def test_object_name(self, widget: InfoArea) -> None:
        assert widget.objectName() == "infoArea"

    def test_view_model_access(self, widget: InfoArea) -> None:
        assert widget.view_model is not None
        assert isinstance(widget.view_model, InfoAreaViewModel)

    def test_status_meta_label_exists(self, widget: InfoArea) -> None:
        assert widget._status_meta_label is not None
        assert widget._status_meta_label.objectName() == "panelCardTitle"

    def test_status_summary_label_exists(self, widget: InfoArea) -> None:
        assert widget._status_summary_label is not None
        assert widget._status_summary_label.objectName() == "infoStatusSummary"

    def test_guide_row_initially_hidden(self, widget: InfoArea) -> None:
        assert widget.is_guide_row_visible is False


class TestActivityEntry:
    def test_entry_is_hidden_until_records_exist(self, widget, vm, qtbot):
        assert widget._activity_button.isHidden()
        vm.set_activity_counts(3, 2)
        assert not widget._activity_button.isHidden()
        assert widget._activity_button.property("hasFailures") is True
        with qtbot.waitSignal(vm.activity_requested):
            widget._activity_button.click()
        vm.set_activity_counts(1, 0)
        assert widget._activity_button.property("hasFailures") is False


class TestStatusSection:
    def test_status_shows_history_summary(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("History entry", "info")
        assert widget.status_summary_text == "History entry"
        assert widget.status_source == "history"

    def test_status_shows_transient(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.set_transient_message("error", "Error!", "danger", ttl_ms=0)
        assert widget.status_summary_text == "Error!"
        assert widget.status_tone == "danger"
        assert widget.status_source == "transient"

    def test_status_shows_task_summary(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(
            operation_id="op-1",
            current_file="file.docx",
            completed_count=1,
            total_count=3,
            failed_count=0,
            state="active",
            tone="info",
        )
        assert widget.status_source == "task"
        assert "file.docx" in widget.status_summary_text

    def test_status_shows_idle_when_empty(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        assert widget.status_source == "idle"
        assert widget._content_card.title() == vm.status_summary_text
        assert not widget._content_card.header.isHidden()
        assert widget._status_summary_label.isHidden()
        assert widget._content_card.header.property("contentCollapsed") is True

    def test_idle_card_expands_for_progress_and_retains_terminal_result(
        self, widget: InfoArea, vm: InfoAreaViewModel
    ) -> None:
        vm.set_task_summary(operation_id="op", current_file="report.md", state="active", tone="info")
        assert widget._content_card.header.property("contentCollapsed") is False
        assert not widget._progress.isHidden()
        assert not widget._status_summary_label.isHidden()
        vm.set_task_summary(
            operation_id="op",
            current_file="report.md",
            state="success",
            tone="success",
            completed_count=1,
            total_count=1,
        )
        assert widget._content_card.header.property("contentCollapsed") is False
        assert widget._progress.isHidden()
        assert widget._content_card.property("panelTone") == "success"
        assert "1/1" not in widget.status_summary_text
        vm.reset_session()
        assert widget._content_card.header.property("contentCollapsed") is True

    def test_idle_card_keeps_activity_and_custom_destination_reachable(
        self, widget: InfoArea, vm: InfoAreaViewModel
    ) -> None:
        vm.set_output_destination_hint("Output: D:/Converted")
        assert widget._content_card.header.property("contentCollapsed") is False
        assert not widget._output_destination_label.isHidden()
        assert widget._status_summary_label.isHidden()
        vm.set_output_destination_hint("")
        assert widget._content_card.header.property("contentCollapsed") is True
        vm.set_activity_counts(2, 1)
        assert widget._content_card.header.property("contentCollapsed") is False
        assert not widget._activity_button.isHidden()

    def test_status_meta_text(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("Test", "info")
        assert len(vm.status_meta_text) > 0


# ── Guide buttons ─────────────────────────────────────────────────────────


class TestGuideButtons:
    def test_renders_guide_buttons_for_success(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        guide_actions = [
            {"action_key": "view_failed_details", "target_path": "/tmp/out"},
        ]
        vm.set_task_summary(
            operation_id="op-2000",
            current_file="done.docx",
            completed_count=2,
            total_count=2,
            failed_count=0,
            state="success",
            tone="success",
            guide_actions=guide_actions,
        )
        assert widget.is_guide_row_visible
        buttons = widget.find_guide_buttons()
        assert len(buttons) == 1
        assert buttons[0].property("guideActionPriority") == "primary"

    def test_renders_guide_buttons_for_failed(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        guide_actions = [
            {"action_key": "view_failed_details", "target_path": "/tmp/failed.txt"},
            {"action_key": "retry_failed", "target_path": ""},
        ]
        vm.set_task_summary(
            operation_id="op-3000",
            current_file="failed.docx",
            completed_count=1,
            total_count=1,
            failed_count=1,
            state="failed",
            tone="danger",
            guide_actions=guide_actions,
        )
        assert widget.is_guide_row_visible
        buttons = widget.find_guide_buttons()
        assert len(buttons) == 2

    def test_guide_not_visible_for_active(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(
            operation_id="op-1",
            state="active",
            tone="info",
        )
        assert not widget.is_guide_row_visible

    def test_guide_button_click_emits(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        emitted: list[tuple[str, str]] = []
        vm.task_guide_action_requested.connect(lambda ak, tp: emitted.append((ak, tp)))
        guide_actions = [
            {"action_key": "view_failed_details", "target_path": "/tmp/out.md"},
        ]
        vm.set_task_summary(
            operation_id="op-1",
            completed_count=1,
            total_count=1,
            state="success",
            tone="success",
            guide_actions=guide_actions,
        )
        buttons = widget.find_guide_buttons()
        assert len(buttons) == 1
        buttons[0].click()
        assert emitted == [("view_failed_details", "/tmp/out.md")]

    def test_guide_buttons_have_minimum_height(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        guide_actions = [
            {"action_key": "view_failed_details", "target_path": "/tmp/out"},
        ]
        vm.set_task_summary(
            operation_id="op-1",
            completed_count=1,
            total_count=1,
            state="success",
            tone="success",
            guide_actions=guide_actions,
        )
        buttons = widget.find_guide_buttons()
        assert buttons[0].minimumHeight() >= 32


# ── ViewModel syncing ─────────────────────────────────────────────────────


class TestViewModelSyncing:
    def test_syncs_on_state_changed(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        """Widget should update display when ViewModel state changes."""
        vm.add_message("First message", "info")
        assert widget.message_count == 1

        vm.add_message("Second message", "success")
        assert widget.message_count == 2

    def test_clear_history_updates_widget(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("msg1", "info")
        vm.add_message("msg2", "info")
        assert widget.message_count == 2
        vm.clear_history()
        assert widget.message_count == 0


# ── Message limit rendering ───────────────────────────────────────────────


class TestMessageLimitRendering:
    def test_widget_reflects_message_limit(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.max_messages = 2
        vm.add_message("info-1", "info")
        vm.add_message("warn-1", "warning")
        vm.add_message("info-2", "info")
        assert widget.message_count == 2
        assert widget.message_types == ["warning", "info"]


# ── Scroll area ───────────────────────────────────────────────────────────
