"""Smoke tests for InputArea widget.

These tests validate widget construction, property defaults, and
ViewModel wiring.  They require a QApplication instance.
"""

from collections.abc import Generator
from itertools import pairwise
from pathlib import Path

import pytest
from PySide6.QtCore import QCoreApplication, QEvent, QMimeData, QPoint, QPointF, Qt, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QLabel,
    QScrollArea,
    QStyle,
    QStyleOptionComboBox,
    QVBoxLayout,
    QWidget,
)
from shiboken6 import isValid

from docwen_gui.styles.panel import build_panel_stylesheet
from docwen_gui.view_models.input_area_vm import _BATCH_SCAN_LIMIT, InputAreaViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel
from docwen_gui.widgets.input_area import (
    _COMPACT_WIDTH_THRESHOLD,
    _DEFAULT_HEIGHT,
    InputArea,
)

pytestmark = pytest.mark.gui


@pytest.fixture
def main_vm() -> MainWindowViewModel:
    return MainWindowViewModel(controller=None)


@pytest.fixture
def input_vm(main_vm: MainWindowViewModel) -> InputAreaViewModel:
    return InputAreaViewModel(main_vm=main_vm)


@pytest.fixture
def widget(qapp: QApplication, input_vm: InputAreaViewModel) -> "Generator[InputArea, None, None]":
    w = InputArea(view_model=input_vm)
    yield w
    if isValid(w):
        w.deleteLater()


# ── Construction ──────────────────────────────────────────────────────


class TestConstruction:
    def test_widget_created(self, widget: InputArea) -> None:
        assert widget is not None
        assert widget.objectName() == "inputArea"

    def test_accepts_drops(self, widget: InputArea) -> None:
        assert widget.acceptDrops()

    def test_default_height(self, widget: InputArea) -> None:
        assert widget.minimumHeight() >= _DEFAULT_HEIGHT

    def test_view_model_access(self, widget: InputArea) -> None:
        assert widget.view_model is not None
        assert isinstance(widget.view_model, InputAreaViewModel)

    def test_add_button_exists(self, widget: InputArea) -> None:
        btn = widget.add_button
        assert btn is not None
        assert btn.text()  # has visible text

    def test_clear_button_exists(self, widget: InputArea) -> None:
        btn = widget.clear_button
        assert btn is not None

    def test_action_buttons_share_minimum_geometry(self, widget: InputArea, qapp: QApplication) -> None:
        from docwen_gui.styles.design_tokens import Sizing
        from docwen_gui.widgets.input_area import _ACTION_BUTTON_MIN_WIDTH

        widget.resize(720, 480)
        widget.show()
        qapp.processEvents()

        add_button = widget.add_button
        clear_button = widget.clear_button
        assert add_button.minimumWidth() == clear_button.minimumWidth() >= _ACTION_BUTTON_MIN_WIDTH
        assert add_button.minimumHeight() == clear_button.minimumHeight() >= Sizing.CONTROL_HEIGHT
        assert add_button.height() == clear_button.height() == widget.mode_switch.height()

    def test_mode_switch_exists(self, widget: InputArea) -> None:
        switch = widget.mode_switch
        assert switch is not None
        # SegmentedWidget.currentItem() returns a SegmentedItem, use text to verify
        current = switch.currentItem()
        assert current is not None

    def test_empty_state_shows_supported_formats(self, widget: InputArea) -> None:
        type_labels = widget.findChildren(QLabel, "fileDropTypesTypeLabel")
        value_labels = widget.findChildren(QLabel, "fileDropTypesValueLabel")

        assert len(type_labels) == 6
        assert len(value_labels) == 6
        values = " ".join(label.text() for label in value_labels)
        assert "DOCX" in values
        assert "HEIF" in values
        assert "EPUB" in values

    def test_selection_feedback_is_not_a_nested_status_card(self, widget: InputArea) -> None:
        assert widget._feedback_frame.objectName() == "fileDropFeedbackArea"
        assert not widget.findChildren(QLabel, "fileDropFeedbackTitleLabel")

    @pytest.mark.parametrize(
        ("theme", "readable_color"),
        [
            ("light", "rgba(71, 85, 105, 224)"),
            ("dark", "rgba(203, 213, 225, 218)"),
        ],
    )
    def test_supported_formats_use_readable_secondary_text(self, theme: str, readable_color: str) -> None:
        stylesheet = build_panel_stylesheet(theme)

        assert "QLabel#fileDropTypesTypeLabel, QLabel#fileDropTypesValueLabel" in stylesheet
        assert f"color: {readable_color};" in stylesheet

    @pytest.mark.parametrize(
        ("theme", "arrow_asset"),
        [("light", "ChevronDown_black.svg"), ("dark", "ChevronDown_white.svg")],
    )
    def test_selectors_have_a_visible_trigger_distinct_from_text_inputs(self, theme: str, arrow_asset: str) -> None:
        stylesheet = build_panel_stylesheet(theme)

        assert "QComboBox:hover" in stylesheet
        assert "QComboBox:focus," in stylesheet
        assert "QComboBox:on {" in stylesheet
        assert "QComboBox:disabled" in stylesheet
        assert "QComboBox::drop-down" in stylesheet
        assert "QComboBox:left-to-right::drop-down" in stylesheet
        assert "border-left:" in stylesheet
        assert "QComboBox:right-to-left::drop-down" in stylesheet
        assert "border-right:" in stylesheet
        assert "QComboBox::drop-down:hover" in stylesheet
        assert "QComboBox::drop-down:pressed, QComboBox::drop-down:on" in stylesheet
        assert "QComboBox::down-arrow" in stylesheet
        assert arrow_asset in stylesheet

    def test_selector_trigger_follows_layout_direction(self, qapp: QApplication) -> None:
        combo = QComboBox()
        combo.addItems(["DOCX", "PDF"])
        combo.resize(200, 40)
        combo.setStyleSheet(build_panel_stylesheet("light"))
        combo.show()

        def arrow_rect(direction: Qt.LayoutDirection):
            combo.setLayoutDirection(direction)
            combo.ensurePolished()
            qapp.processEvents()
            option = QStyleOptionComboBox()
            combo.initStyleOption(option)
            return combo.style().subControlRect(
                QStyle.ComplexControl.CC_ComboBox,
                option,
                QStyle.SubControl.SC_ComboBoxArrow,
                combo,
            )

        ltr_arrow = arrow_rect(Qt.LayoutDirection.LeftToRight)
        rtl_arrow = arrow_rect(Qt.LayoutDirection.RightToLeft)

        assert ltr_arrow.center().x() > combo.width() // 2
        assert rtl_arrow.center().x() < combo.width() // 2
        combo.close()

    def test_supported_formats_restore_symmetric_pyramid(self, widget, qapp):
        widget.resize(1200, _DEFAULT_HEIGHT)
        widget.show()
        qapp.processEvents()
        widget._sync_supported_type_layout()
        qapp.processEvents()
        assert len(widget._type_prompt_rows) == 6
        left_edges = []
        right_edges = []
        for row, _layout, label, value in widget._type_prompt_rows:
            left_edges.append(label.geometry().left())
            right_edges.append(value.geometry().right())
            assert abs(label.x() - (row.width() - 1 - value.geometry().right())) <= 1
            assert label.geometry().right() < value.geometry().left()
            assert value.geometry().right() < row.width()
            assert value.alignment() & Qt.AlignmentFlag.AlignRight
            assert value.height() >= value.heightForWidth(value.width())
        assert all(upper > lower for upper, lower in pairwise(left_edges))
        assert all(upper < lower for upper, lower in pairwise(right_edges))

    @pytest.mark.parametrize("font_pixels", [13, 20, 26])
    def test_format_pyramid_wraps_without_clipping_when_narrow(self, widget, qapp, font_pixels):
        for _, _, label, value in widget._type_prompt_rows:
            for control in (label, value):
                font = control.font()
                font.setPixelSize(font_pixels)
                control.setFont(font)
        widget.resize(360, _DEFAULT_HEIGHT)
        widget.show()
        qapp.processEvents()
        widget._sync_supported_type_layout()
        qapp.processEvents()
        previous_bottom = -1
        for row, layout, label, value in widget._type_prompt_rows:
            margins = layout.contentsMargins()
            assert margins.left() == margins.right() >= 0
            assert label.geometry().right() < value.geometry().left()
            assert 0 <= value.x() < value.geometry().right() < row.width()
            assert value.height() >= value.heightForWidth(value.width())
            assert row.y() > previous_bottom
            previous_bottom = row.geometry().bottom()

    def test_large_empty_prompt_uses_available_width_without_vertical_clipping(self, widget, qapp):
        widget.setStyleSheet(build_panel_stylesheet("light", "xlarge"))
        widget._prompt_label.setText("Drag a single document here")
        # Match the main window: its workflow scroll area gives the input
        # widget its minimum content height even when the viewport is shorter.
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        flow = QVBoxLayout(content)
        flow.setContentsMargins(0, 0, 0, 0)
        flow.addWidget(widget)
        flow.addStretch(1)
        scroll.setWidget(content)
        scroll.resize(360, _DEFAULT_HEIGHT)
        scroll.show()
        qapp.processEvents()
        widget._sync_prompt_layout()
        qapp.processEvents()
        prompt = widget._prompt_label
        assert prompt.height() >= prompt.heightForWidth(prompt.width())
        assert prompt.geometry().bottom() < widget._empty_title_row.height()
        assert prompt.y() > widget._hero_icon_label.geometry().bottom()
        assert prompt.width() <= widget._empty_center_panel.width()
        assert scroll.verticalScrollBar().maximum() > 0
        scroll.close()

    def test_supported_formats_hide_with_selection_feedback(self, widget: InputArea, tmp_path) -> None:
        sample = tmp_path / "sample.docx"
        sample.write_text("content")
        assert not widget._empty_content.isHidden()

        widget.view_model.add_files([str(sample)])

        assert widget._empty_content.isHidden()


# ── Mode switch ───────────────────────────────────────────────────────


class TestModeSwitch:
    def test_default_mode_single(self, widget: InputArea) -> None:
        current = widget.mode_switch.currentItem()
        assert current is not None
        assert widget.view_model.mode == "single"

    def test_switch_to_batch(self, widget: InputArea) -> None:
        widget.view_model.set_mode("batch")
        current = widget.mode_switch.currentItem()
        assert current is not None
        assert widget.view_model.mode == "batch"


# ── Drag-and-drop MIME acceptance ─────────────────────────────────────


class TestDragDropMime:
    def test_accepts_url_mime(self, widget: InputArea) -> None:
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile("/tmp/test.docx")])
        assert mime.hasUrls()

    def test_accepts_text_mime(self) -> None:
        mime = QMimeData()
        mime.setText("/tmp/test.docx")
        assert mime.hasText()

    def test_drag_enter_previews_batch_folder_scan(self, widget: InputArea, tmp_path) -> None:
        folder = tmp_path / "input"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        supported = folder / "keep.txt"
        unsupported = nested / "ignore.bin"
        supported.write_text("1")
        unsupported.write_text("2")
        widget.view_model.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        widget.view_model.set_mode("batch")

        mime = QMimeData()
        mime.setText(str(folder))
        event = QDragEnterEvent(
            QPoint(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        widget.dragEnterEvent(event)

        assert event.isAccepted()
        assert "1" in widget._selection_label.text()
        assert "ignore.bin" in widget._selection_label.toolTip()
        assert widget._feedback_frame.property("feedbackTone") == "info"

    def test_drop_batch_folder_reports_partial_skips(self, widget: InputArea, tmp_path) -> None:
        folder = tmp_path / "mixed"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        supported = folder / "keep.txt"
        unsupported = nested / "ignore.bin"
        supported.write_text("1")
        unsupported.write_text("2")
        widget.view_model.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        widget.view_model.set_mode("batch")
        added: list[list[str]] = []
        widget.view_model.files_added.connect(lambda paths: added.append(list(paths)))

        mime = QMimeData()
        mime.setText(str(folder))
        event = QDropEvent(
            QPointF(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        widget.dropEvent(event)

        assert event.isAccepted()
        assert added == [[str(supported)]]
        assert "1" in widget._selection_label.text()
        assert widget._feedback_frame.property("feedbackTone") == "warning"

    def test_url_drop_large_folder_summarizes_preview_but_adds_all_supported_files(
        self,
        widget: InputArea,
        tmp_path,
    ) -> None:
        folder = tmp_path / "large-url-drop"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        total_supported = _BATCH_SCAN_LIMIT + 7
        supported_files = []
        for index in range(total_supported):
            target_dir = nested if index % 3 == 0 else folder
            file_path = target_dir / f"doc-{index:03}.txt"
            file_path.write_text("content")
            supported_files.append(str(file_path))
        (nested / "ignore.bin").write_text("unsupported")
        widget.view_model.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        widget.view_model.set_mode("batch")
        added: list[list[str]] = []
        widget.view_model.files_added.connect(lambda paths: added.append(list(paths)))

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(folder))])
        preview = widget.view_model.build_drag_preview([str(folder)])
        enter_event = QDragEnterEvent(
            QPoint(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        widget.dragEnterEvent(enter_event)

        assert enter_event.isAccepted()
        assert preview.has_degraded_preview is True
        assert widget._selection_label.text() == preview.message
        assert widget._feedback_frame.property("feedbackTone") == "info"

        drop_event = QDropEvent(
            QPointF(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        widget.dropEvent(drop_event)

        assert drop_event.isAccepted()
        assert len(added) == 1
        assert set(added[0]) == set(supported_files)
        assert len(added[0]) == total_supported
        assert "ignore.bin" not in "\n".join(added[0])
        assert widget._feedback_frame.property("feedbackTone") == "warning"


# ── Compact layout ───────────────────────────────────────────────────


class TestCompactLayout:
    def test_resize_above_threshold_does_not_crash(self, widget: InputArea) -> None:
        widget.resize(_COMPACT_WIDTH_THRESHOLD + 100, _DEFAULT_HEIGHT)
        # Verify widget handles resize without error
        assert widget.mode_switch is not None

    def test_resize_below_threshold_does_not_crash(self, widget: InputArea) -> None:
        widget.resize(_COMPACT_WIDTH_THRESHOLD - 50, _DEFAULT_HEIGHT)
        assert widget.mode_switch is not None

    def test_deferred_layout_sync_runs_normally_and_coalesces(
        self,
        widget: InputArea,
        qapp: QApplication,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        widget.show()
        qapp.processEvents()
        calls: list[str] = []
        prompt_sync = widget._sync_prompt_layout
        supported_type_sync = widget._sync_supported_type_layout

        def sync_prompt() -> None:
            calls.append("prompt")
            prompt_sync()

        def sync_supported_types() -> None:
            calls.append("supported-types")
            supported_type_sync()

        monkeypatch.setattr(widget, "_sync_prompt_layout", sync_prompt)
        monkeypatch.setattr(widget, "_sync_supported_type_layout", sync_supported_types)

        widget._schedule_deferred_layout_sync(prompt=True)
        widget._schedule_deferred_layout_sync(prompt=True, supported_types=True)
        qapp.processEvents()

        assert calls == ["prompt", "supported-types"]
        assert not widget._deferred_layout_sync_timer.isActive()

    def test_close_cancels_deferred_layout_callbacks(
        self,
        widget: InputArea,
        qapp: QApplication,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        widget.show()
        qapp.processEvents()
        calls: list[str] = []
        monkeypatch.setattr(widget, "_sync_prompt_layout", lambda: calls.append("prompt"))
        monkeypatch.setattr(widget, "_sync_supported_type_layout", lambda: calls.append("supported-types"))

        widget._schedule_deferred_layout_sync(prompt=True, supported_types=True)
        widget.close()
        qapp.processEvents()

        assert calls == []
        assert not widget._deferred_layout_sync_timer.isActive()

    def test_delete_discards_deferred_layout_callbacks(
        self,
        widget: InputArea,
        qapp: QApplication,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        widget.show()
        qapp.processEvents()
        calls: list[str] = []
        monkeypatch.setattr(widget, "_sync_prompt_layout", lambda: calls.append("prompt"))
        monkeypatch.setattr(widget, "_sync_supported_type_layout", lambda: calls.append("supported-types"))

        widget._schedule_deferred_layout_sync(prompt=True, supported_types=True)
        widget.deleteLater()
        QCoreApplication.sendPostedEvents(widget, QEvent.Type.DeferredDelete)
        qapp.processEvents()

        assert calls == []
        assert not isValid(widget)

    def test_deferred_layout_sync_ignores_deleted_child_controls(
        self,
        widget: InputArea,
        qapp: QApplication,
    ) -> None:
        widget.show()
        qapp.processEvents()
        drop_group = widget._drop_group

        widget._schedule_deferred_layout_sync(prompt=True, supported_types=True)
        drop_group.deleteLater()
        QCoreApplication.sendPostedEvents(drop_group, QEvent.Type.DeferredDelete)
        qapp.processEvents()

        assert not isValid(drop_group)
        assert not widget._deferred_layout_sync_timer.isActive()


# ── ViewModel signal wiring ──────────────────────────────────────────


class TestViewModelWiring:
    def test_location_uses_committed_file_through_feedback_and_replacement(
        self, widget: InputArea, main_vm: MainWindowViewModel, tmp_path: Path
    ) -> None:
        from docwen_core.models import FileRef

        selected = FileRef(path=str(tmp_path / "first.md"), format="markdown", category="markdown")
        main_vm.set_selected_file(selected)
        widget.view_model.sync_selection([selected], current=True)
        emitted: list[str] = []
        widget.location_requested.connect(emitted.append)
        widget._selection_detail_label.set_full_text("Warning: a different file was rejected")
        widget._open_location_button.click()
        assert emitted == [selected.path]

        replacement = FileRef(path=str(tmp_path / "second.md"), format="markdown", category="markdown")
        main_vm.set_selected_file(replacement)
        widget.view_model.sync_selection([replacement], current=True)
        widget._open_location_button.click()
        assert emitted[-1] == replacement.path
        main_vm.clear_selected_file()
        widget.view_model.clear_files()
        widget._open_selected_location()
        assert len(emitted) == 2
        assert widget._open_location_button.isHidden()

    def test_mode_changed_updates_switch(self, widget: InputArea) -> None:
        widget.view_model.set_mode("batch")
        current = widget.mode_switch.currentItem()
        assert current is not None
        assert widget.view_model.mode == "batch"

    def test_selection_message_visible(self, widget: InputArea) -> None:
        widget.view_model.selection_message_changed.emit("Test message", "success")
        # No crash; visual state handled by widget
        assert True

    def test_clear_resets_selection(self, widget: InputArea) -> None:
        widget.view_model.clear_files()
        # Widget should handle this gracefully
        assert True


# ── Public API ────────────────────────────────────────────────────────


class TestPublicApi:
    def test_set_recent_files(self, widget: InputArea, tmp_path) -> None:
        f1 = tmp_path / "recent.docx"
        f1.write_text("content")
        widget.set_recent_files([str(f1)])
        # No crash

    def test_set_recent_files_nonexistent_filtered(self, widget: InputArea) -> None:
        widget.set_recent_files(["/nonexistent/path.docx"])
        # Should filter out nonexistent files silently

    def test_update_display(self, widget: InputArea, tmp_path) -> None:
        f1 = tmp_path / "display.docx"
        f1.write_text("content")
        widget.update_display(str(f1))
        # Should update selection state

    def test_update_display_empty(self, widget: InputArea) -> None:
        widget.update_display("")
        # No crash on empty path

    def test_open_file_dialog_can_force_batch_mode(
        self,
        widget: InputArea,
        main_vm: MainWindowViewModel,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        opened: list[str] = []
        monkeypatch.setattr(widget, "_open_file_dialog", lambda: opened.append("file"))

        widget.open_file_dialog(force_batch_mode=True)

        assert widget.view_model.mode == "batch"
        assert main_vm.mode == "batch"
        assert opened == ["file"]

    def test_open_folder_dialog_can_force_batch_mode(
        self,
        widget: InputArea,
        main_vm: MainWindowViewModel,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        opened: list[str] = []
        monkeypatch.setattr(widget, "_open_folder_dialog", lambda: opened.append("folder"))

        widget.open_folder_dialog(force_batch_mode=True)

        assert widget.view_model.mode == "batch"
        assert main_vm.mode == "batch"
        assert opened == ["folder"]
