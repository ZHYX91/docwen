"""Focused tests split from test_batch_list_widget.py."""

from __future__ import annotations

import pytest

from ._batch_list_widget_support import (
    BatchEntryItemWidget,
    BatchList,
    BatchListViewModel,
    QApplication,
    QEvent,
    QKeyEvent,
    QListWidgetItem,
    Qt,
    ReorderableListWidget,
    WrapRowLayout,
    _add_synthetic,
    _t,
)
from ._batch_list_widget_support import (
    populated_widget as populated_widget,
)

pytestmark = pytest.mark.gui
from ._batch_list_widget_support import (
    vm as vm,
)
from ._batch_list_widget_support import (
    widget as widget,
)


class TestFilter:
    def test_filter_button_exists(self, populated_widget: BatchList) -> None:
        assert populated_widget.filter_button is not None

    def test_set_filter_via_viewmodel(self, populated_widget: BatchList) -> None:
        vm = populated_widget.view_model
        vm.set_status_filter("failed")
        assert vm.active_filter == "failed"

    def test_filter_moves_current_file_to_visible_entry(self, qapp: QApplication) -> None:
        vm = BatchListViewModel()
        pending = "/test/pending.docx"
        failed = "/test/failed.docx"
        _add_synthetic(vm, [pending, failed])
        vm.set_file_status(failed, "failed", error_message="boom")
        widget = BatchList(view_model=vm)
        try:
            vm.activate_tab("document")
            list_widget = widget._tabs["document"]
            list_widget.setCurrentRow(0)
            assert widget.get_current_file() == pending

            vm.set_status_filter("failed")

            assert widget.get_current_file() == failed
            assert list_widget.currentItem() is not None
            assert not list_widget.currentItem().isHidden()
        finally:
            widget.deleteLater()

    def test_reorder_buttons_respect_visible_filter(self, qapp: QApplication) -> None:
        vm = BatchListViewModel()
        pending = "/test/pending.docx"
        failed = "/test/failed.docx"
        _add_synthetic(vm, [pending, failed])
        vm.set_file_status(failed, "failed", error_message="boom")
        widget = BatchList(view_model=vm)
        try:
            vm.activate_tab("document")
            list_widget = widget._tabs["document"]
            list_widget.setCurrentRow(0)

            vm.set_status_filter("failed")
            widget._update_reorder_buttons()

            assert widget.get_current_file() == failed
            assert not widget.move_up_button.isEnabled()
            assert not widget.move_down_button.isEnabled()
        finally:
            widget.deleteLater()


class TestSortWidget:
    def test_sort_button_exists(self, populated_widget: BatchList) -> None:
        assert populated_widget.sort_button is not None

    def test_set_sort_via_viewmodel(self, populated_widget: BatchList) -> None:
        vm = populated_widget.view_model
        vm.set_sort_state("name", False)
        assert vm.sort_key == "name"
        assert vm.sort_ascending is False

    def test_sort_button_updates_after_manual_reorder(self, qapp: QApplication) -> None:
        vm = BatchListViewModel()
        _add_synthetic(vm, ["/test/a.docx", "/test/b.docx"])
        widget = BatchList(view_model=vm)
        try:
            vm.activate_tab("document")
            list_widget = widget._tabs["document"]

            vm.set_sort_state("name", False)
            assert _t("components.file_drop.batch_list.sort_name") in widget.sort_button.text()

            list_widget.setCurrentRow(0)
            assert widget._move_current_item_down() is True

            assert vm.sort_key == "custom"
            assert _t("components.file_drop.batch_list.sort_custom") in widget.sort_button.text()
        finally:
            widget.deleteLater()


class TestMoveReorder:
    def test_reorder_buttons_expose_keyboard_shortcuts(self, populated_widget: BatchList) -> None:
        assert "Ctrl+↑" in populated_widget.move_up_button.toolTip()
        assert "Ctrl+↓" in populated_widget.move_down_button.toolTip()
        assert populated_widget.move_up_button.accessibleDescription() == populated_widget.move_up_button.toolTip()
        assert populated_widget.move_down_button.accessibleDescription() == populated_widget.move_down_button.toolTip()

    def test_move_up_disabled_for_first_item(self, populated_widget: BatchList) -> None:
        # File order: doc.docx, sheet.xlsx, img.png, layout.pdf, readme.md, book.epub
        # Current tab is text (after adding, first tab got focus)
        list_widget = populated_widget._tabs[populated_widget._vm.current_category]
        if list_widget.count() >= 2:
            list_widget.setCurrentRow(0)
            populated_widget._update_reorder_buttons()
            assert not populated_widget.move_up_button.isEnabled()
            assert populated_widget.move_down_button.isEnabled() or list_widget.count() <= 1

    def test_move_down_disabled_for_last_item(self, populated_widget: BatchList) -> None:
        list_widget = populated_widget._tabs[populated_widget._vm.current_category]
        count = list_widget.count()
        if count >= 2:
            list_widget.setCurrentRow(count - 1)
            populated_widget._update_reorder_buttons()
            assert populated_widget.move_up_button.isEnabled()
            assert not populated_widget.move_down_button.isEnabled()

    def test_ctrl_up_handled(self, qapp: QApplication) -> None:
        """ReorderableListWidget handles Ctrl+Up key press."""
        rw = ReorderableListWidget()
        item1 = QListWidgetItem("a")
        item2 = QListWidgetItem("b")
        rw.addItem(item1)
        rw.addItem(item2)
        rw.setCurrentRow(1)  # select "b"

        rw.item_reordered.connect(lambda: setattr(self, "_reordered", True))
        self._reordered = False

        event = QKeyEvent(
            QEvent.Type.KeyPress,
            Qt.Key.Key_Up,
            Qt.KeyboardModifier.ControlModifier,
        )
        rw.keyPressEvent(event)
        assert rw.currentRow() == 0  # moved up
        assert rw.item(0).text() == "b"
        assert rw.item(1).text() == "a"
        rw.deleteLater()

    def test_ctrl_down_handled(self, qapp: QApplication) -> None:
        rw = ReorderableListWidget()
        item1 = QListWidgetItem("a")
        item2 = QListWidgetItem("b")
        rw.addItem(item1)
        rw.addItem(item2)
        rw.setCurrentRow(0)  # select "a"

        event = QKeyEvent(
            QEvent.Type.KeyPress,
            Qt.Key.Key_Down,
            Qt.KeyboardModifier.ControlModifier,
        )
        rw.keyPressEvent(event)
        assert rw.currentRow() == 1  # moved down
        assert rw.item(1).text() == "a"
        assert rw.item(0).text() == "b"
        rw.deleteLater()

    def test_manual_reorder_refreshes_sequence_markers(self, qapp: QApplication) -> None:
        vm = BatchListViewModel()
        _add_synthetic(vm, ["/test/a.docx", "/test/b.docx"])
        widget = BatchList(view_model=vm)
        try:
            vm.activate_tab("document")
            list_widget = widget._tabs["document"]
            list_widget.setCurrentRow(0)

            assert widget._move_current_item_down() is True

            first = list_widget.itemWidget(list_widget.item(0))
            second = list_widget.itemWidget(list_widget.item(1))
            assert isinstance(first, BatchEntryItemWidget)
            assert isinstance(second, BatchEntryItemWidget)
            assert first.status_icon_label.text() == "1"
            assert second.status_icon_label.text() == "2"
        finally:
            widget.deleteLater()

    def test_ctrl_up_at_top_no_op(self, qapp: QApplication) -> None:
        rw = ReorderableListWidget()
        rw.addItem(QListWidgetItem("a"))
        rw.setCurrentRow(0)
        event = QKeyEvent(
            QEvent.Type.KeyPress,
            Qt.Key.Key_Up,
            Qt.KeyboardModifier.ControlModifier,
        )
        rw.keyPressEvent(event)
        assert rw.currentRow() == 0  # unchanged
        rw.deleteLater()


class TestContextMenu:
    def test_build_retry_targets(self, populated_widget: BatchList) -> None:
        """Verify retry target building does not crash."""
        category = populated_widget._vm.current_category
        selected, category_failed = populated_widget._vm.build_retry_targets(category, None)
        assert isinstance(selected, list)
        assert isinstance(category_failed, list)


class TestSelection:
    def test_selection_changed_signal(self, populated_widget: BatchList) -> None:
        emitted: list[object] = []
        populated_widget.selection_changed.connect(lambda f: emitted.append(f))
        # Trigger selection change via current item change
        list_widget = populated_widget._tabs[populated_widget._vm.current_category]
        if list_widget.count() > 0:
            list_widget.setCurrentRow(0)

    def test_select_file_accepts_raw_windows_path(self, qapp: QApplication, vm: BatchListViewModel) -> None:
        stored_path = "C:/Users/Example/Documents/report.docx"
        raw_path = r"C:\Users\Example\Documents\report.docx"
        _add_synthetic(vm, [stored_path])
        w = BatchList(view_model=vm)
        try:
            assert w.select_file(raw_path) is True
            assert w.get_current_file() == stored_path
        finally:
            w.deleteLater()


class TestEntryActionSignal:
    def test_entry_action_requested_emitted(self, populated_widget: BatchList) -> None:
        emitted: list[tuple] = []
        populated_widget.entry_action_requested.connect(lambda k, p: emitted.append((k, p)))
        populated_widget._handle_entry_action("test_key", "/test/path")
        assert len(emitted) == 1
        assert emitted[0] == ("test_key", "/test/path")

    def test_retry_failed_handled_internally(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/unique_file.docx"])
        vm.set_file_status("/test/unique_file.docx", "failed", error_message="err")
        # reset_failed_files is called via the widget's handler
        vm.reset_failed_files(["/test/unique_file.docx"])
        entry = vm.get_file_entry("/test/unique_file.docx")
        assert entry is not None
        assert entry.status == "pending"

    def test_remove_entry_handled_internally(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/unique_file.docx"])
        assert vm.remove_file("/test/unique_file.docx") is True
        assert vm.entry_count == 0


class TestSummaryRefresh:
    def test_summary_after_add(self, populated_widget: BatchList) -> None:
        assert len(populated_widget.summary_label.text()) > 0

    def test_has_files_property_set(self, populated_widget: BatchList) -> None:
        assert populated_widget.summary_section.property("hasFiles") is True


class TestReorderableListWidget:
    def test_move_current_item_by(self, qapp: QApplication) -> None:
        rw = ReorderableListWidget()
        for ch in "abcde":
            rw.addItem(QListWidgetItem(ch))
        rw.setCurrentRow(2)  # 'c'
        rw.move_current_item_by(-2)  # move to position 0
        assert rw.currentRow() == 0
        assert rw.item(0).text() == "c"
        rw.deleteLater()

    def test_move_current_item_by_uses_visible_neighbors(self, qapp: QApplication) -> None:
        rw = ReorderableListWidget()
        item_a = QListWidgetItem("a")
        item_hidden = QListWidgetItem("hidden")
        item_b = QListWidgetItem("b")
        rw.addItem(item_a)
        rw.addItem(item_hidden)
        rw.addItem(item_b)
        item_hidden.setHidden(True)

        rw.setCurrentItem(item_a)

        assert rw.move_current_item_by(-1) is False
        assert rw.move_current_item_by(1) is True
        assert [rw.item(i).text() for i in range(rw.count()) if not rw.item(i).isHidden()] == ["b", "a"]
        assert rw.currentItem().text() == "a"

        assert rw.move_current_item_by(-1) is True
        assert [rw.item(i).text() for i in range(rw.count()) if not rw.item(i).isHidden()] == ["a", "b"]
        rw.deleteLater()

    def test_move_beyond_bounds_returns_false(self, qapp: QApplication) -> None:
        rw = ReorderableListWidget()
        rw.addItem(QListWidgetItem("a"))
        rw.setCurrentRow(0)
        assert rw.move_current_item_by(-1) is False
        assert rw.move_current_item_by(1) is False
        rw.deleteLater()

    def test_no_selection_returns_false(self, qapp: QApplication) -> None:
        rw = ReorderableListWidget()
        rw.addItem(QListWidgetItem("a"))
        assert rw.move_current_item_by(1) is False
        rw.deleteLater()


class TestWrapRowLayout:
    def test_create_layout(self, qapp: QApplication) -> None:
        from PySide6.QtWidgets import QWidget

        container = QWidget()
        layout = WrapRowLayout(container, spacing=4)
        assert layout.count() == 0
        assert layout.hasHeightForWidth() is True
        container.deleteLater()

    def test_add_items(self, qapp: QApplication) -> None:
        from PySide6.QtWidgets import QLabel, QWidget

        container = QWidget()
        layout = WrapRowLayout(container, spacing=4)
        for text in ["A", "BB", "CCC"]:
            layout.addWidget(QLabel(text))
        assert layout.count() == 3
        container.deleteLater()


class TestFocus:
    def test_focus_in_event(self, populated_widget: BatchList) -> None:
        from PySide6.QtGui import QFocusEvent

        event = QFocusEvent(QEvent.Type.FocusIn, Qt.FocusReason.TabFocusReason)
        populated_widget.focusInEvent(event)


def test_batch_list_importable() -> None:
    from docwen_gui.widgets.batch_list import BatchList as BL

    assert BL is not None
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel as BLVM

    assert BLVM is not None
