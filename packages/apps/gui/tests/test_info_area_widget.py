"""Smoke tests for InfoArea widget.

Tests widget construction, history rendering, transient display,
task summary, guide buttons, and location button interaction.
Requires a QApplication instance.
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QLabel,
    QScrollArea,
    QSizePolicy,
    QWidget,
)

from docwen_gui.view_models.info_area_vm import (
    InfoAreaViewModel,
)
from docwen_gui.widgets.info_area import InfoArea, _StatusLocationButton

pytestmark = pytest.mark.gui


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

    def test_scroll_area_exists(self, widget: InfoArea) -> None:
        assert widget._scroll is not None
        assert widget._scroll.objectName() == "infoHistoryScrollArea"

    def test_status_meta_label_exists(self, widget: InfoArea) -> None:
        assert widget._status_meta_label is not None
        assert widget._status_meta_label.objectName() == "infoStatusMeta"

    def test_status_summary_label_exists(self, widget: InfoArea) -> None:
        assert widget._status_summary_label is not None
        assert widget._status_summary_label.objectName() == "infoStatusSummary"

    def test_guide_row_initially_hidden(self, widget: InfoArea) -> None:
        assert widget.is_guide_row_visible is False

    def test_location_button_context_menu_copies_exact_path(
        self,
        qapp: QApplication,
    ) -> None:
        button = _StatusLocationButton("/test/output/report.docx")
        try:
            menu = button._create_context_menu()
            actions = menu.actions()
            actions[0].trigger()
            assert actions
            assert QApplication.clipboard().text() == "/test/output/report.docx"
        finally:
            button.deleteLater()


# ── History rendering ─────────────────────────────────────────────────────


class TestHistoryRendering:
    def test_history_scroll_area_does_not_raise_the_parent_minimum_height(self, widget: InfoArea) -> None:
        scroll = widget.findChild(QScrollArea, "infoHistoryScrollArea")
        assert scroll is not None
        assert scroll.sizePolicy().verticalPolicy() == QSizePolicy.Policy.Ignored

    def test_empty_history_has_a_centered_semantic_state(self, widget: InfoArea) -> None:
        empty_state = widget.findChild(QWidget, "infoHistoryEmptyState")
        assert empty_state is not None
        assert empty_state.isVisibleTo(widget)

        title = empty_state.findChild(QLabel, "infoHistoryEmptyTitle")
        caption = empty_state.findChild(QLabel, "infoHistoryEmptyCaption")
        assert title is not None and title.text()
        assert caption is not None and caption.text()
        assert title.alignment() & Qt.AlignmentFlag.AlignHCenter
        assert caption.alignment() & Qt.AlignmentFlag.AlignHCenter

    def test_renders_history_row(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("Need attention", "warning")
        assert widget.message_count == 1

        empty_state = widget.findChild(QWidget, "infoHistoryEmptyState")
        assert empty_state is not None
        assert not empty_state.isVisible()

        row = widget.get_history_row_widget(0)
        assert row is not None
        assert row.property("infoStatusTone") == "warning"

    def test_renders_timestamp_label(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("Test message", "info")
        row = widget.get_history_row_widget(0)
        assert row is not None

        timestamps = [lbl for lbl in row.findChildren(QLabel) if lbl.objectName() == "statusTimestamp"]
        assert len(timestamps) == 1
        # Fixed width
        assert timestamps[0].minimumWidth() == timestamps[0].maximumWidth()
        assert timestamps[0].minimumWidth() > 0

    def test_renders_message_text(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("Hello world", "success")
        row = widget.get_history_row_widget(0)
        assert row is not None

        messages = [lbl for lbl in row.findChildren(QLabel) if lbl.objectName() == "infoHistoryText"]
        assert len(messages) == 1
        assert messages[0].text() == "Hello world"
        assert messages[0].toolTip() == "Hello world"
        assert messages[0].textInteractionFlags() == Qt.TextInteractionFlag.TextSelectableByMouse
        assert row.toolTip() == "Hello world"

    def test_multiple_rows(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("First", "info")
        vm.add_message("Second", "success")
        vm.add_message("Third", "warning")
        assert widget.message_count == 3

    def test_rebuild_detaches_stale_history_rows(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("First", "info")
        vm.add_message("Second", "success")
        vm.add_message("Third", "warning")

        rows = widget._msg_container.findChildren(QWidget, "infoHistoryRow")
        assert len(rows) == widget.message_count
        assert [row.parentWidget() for row in rows] == [widget._msg_container] * widget.message_count

    def test_no_location_button_without_show_location(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("No location", "info")
        buttons = widget.find_location_buttons()
        assert len(buttons) == 0


# ── Location button ──────────────────────────────────────────────────────


class TestLocationButton:
    def test_shows_location_button(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("Locate me", "success", show_location=True, file_path="/tmp/test.txt")
        buttons = widget.find_location_buttons()
        assert len(buttons) == 1
        assert buttons[0].objectName() == "statusLocationButton"
        assert buttons[0].toolButtonStyle() == Qt.ToolButtonStyle.ToolButtonIconOnly

    def test_location_button_tooltip(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("Locate me", "success", show_location=True, file_path="/tmp/test.txt")
        buttons = widget.find_location_buttons()
        assert len(buttons) == 1
        assert "/tmp/test.txt" in buttons[0].accessibleDescription()


# ── Status section ────────────────────────────────────────────────────────


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

    def test_status_meta_text(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        vm.add_message("Test", "info")
        assert len(vm.status_meta_text) > 0


# ── Guide buttons ─────────────────────────────────────────────────────────


class TestGuideButtons:
    def test_renders_guide_buttons_for_success(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        guide_actions = [
            {"action_key": "open_output_dir", "target_path": "/tmp/out"},
            {"action_key": "add_more_files", "target_path": ""},
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
        assert len(buttons) == 2
        # First button is primary
        assert buttons[0].property("guideActionPriority") == "primary"
        assert buttons[1].property("guideActionPriority") == "secondary"

    def test_renders_guide_buttons_for_failed(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        guide_actions = [
            {"action_key": "view_failed_details", "target_path": "/tmp/failed.txt"},
            {"action_key": "retry_failed", "target_path": ""},
            {"action_key": "add_more_files", "target_path": ""},
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
        assert len(buttons) == 3

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
            {"action_key": "open_output_dir", "target_path": "/tmp/out.md"},
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
        assert emitted == [("open_output_dir", "/tmp/out.md")]

    def test_guide_buttons_have_minimum_height(self, widget: InfoArea, vm: InfoAreaViewModel) -> None:
        guide_actions = [
            {"action_key": "open_output_dir", "target_path": "/tmp/out"},
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


class TestScrollArea:
    def test_horizontal_scrollbar_always_off(self, widget: InfoArea) -> None:
        assert widget._scroll is not None
        assert widget._scroll.horizontalScrollBarPolicy() == Qt.ScrollBarPolicy.ScrollBarAlwaysOff

    def test_widget_resizable(self, widget: InfoArea) -> None:
        assert widget._scroll is not None
        assert widget._scroll.widgetResizable() is True
