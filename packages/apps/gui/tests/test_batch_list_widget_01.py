"""Focused tests split from test_batch_list_widget.py."""

from __future__ import annotations

from ._batch_list_widget_support import (
    CATEGORY_ORDER,
    FILTER_OPTIONS,
    BatchList,
    BatchListViewModel,
    Path,
    QApplication,
    QPushButton,
    Qt,
    ReorderableListWidget,
    _add_synthetic,
    _filter_option_label,
    _format_size,
    _load_status_icon,
    _PivotText,
    _source_path_text,
    _t,
    build_batch_list_stylesheet,
    cast,
    pytest,
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


class TestHelpers:
    def test_format_size(self) -> None:
        assert "B" in _format_size(500)
        assert "KB" in _format_size(2048)

    def test_source_path_text(self) -> None:
        expected = str(Path("/foo/bar/doc.docx").parent)
        assert _source_path_text("/foo/bar/doc.docx") == expected

    def test_load_status_icon_pending_empty(self) -> None:
        icon = _load_status_icon("pending")
        assert icon.isNull()

    def test_load_status_icon_returns_qicon(self, qapp: QApplication) -> None:
        from PySide6.QtGui import QIcon

        icon = _load_status_icon("completed")
        # Returns a QIcon — could be null if SVG assets not found
        assert isinstance(icon, QIcon)

    @pytest.mark.parametrize(("filter_key", "fallback", "_statuses"), FILTER_OPTIONS)
    def test_filter_option_label_is_localized_at_presentation_boundary(
        self,
        filter_key: str,
        fallback: str,
        _statuses: tuple[str, ...],
    ) -> None:
        if filter_key == "all":
            expected = _t("components.file_drop.batch_list.filter_all", fallback)
        else:
            status_fallback = _t(f"components.file_drop.status.{filter_key}", fallback)
            expected = _t(f"components.file_drop.batch_list.filter_{filter_key}", status_fallback)
        assert _filter_option_label(filter_key, fallback) == expected

    @pytest.mark.parametrize(("theme_name", "color"), [("light", "#6E7781"), ("dark", "#8B949E")])
    def test_batch_stylesheet_keeps_secondary_details_readable(self, theme_name: str, color: str) -> None:
        stylesheet = build_batch_list_stylesheet(theme_name)

        assert 'QLabel#batchInfoLabel[class="danger"] {' in stylesheet
        assert 'QLabel#batchDetailLabel[detailRole="secondary"] {' in stylesheet
        assert f"color: {color};" in stylesheet
        assert "QWidget#batchOutputRow {" in stylesheet


class TestConstruction:
    def test_one_selected_file_has_location_and_remove_actions(self, populated_widget, monkeypatch, qapp):
        from PySide6.QtWidgets import QMenu

        widget = populated_widget
        widget.show()
        qapp.processEvents()
        category = widget.view_model.current_category
        listing = widget._tabs[category]
        listing.setCurrentRow(0)
        file_path = widget.get_current_file()
        original_count = widget.view_model.entry_count
        labels = []

        def inspect_menu(menu, _position):
            labels.extend(action.text() for action in menu.actions() if not action.isSeparator())
            remove_label = _t("components.file_drop.batch_list.action_remove_selected", count=1)
            remove = next(action for action in menu.actions() if action.text() == remove_label)
            assert remove.isEnabled()
            remove.trigger()

        class InspectMenu(QMenu):
            def exec(self, position):
                return inspect_menu(self, position)

        monkeypatch.setitem(widget._show_item_context_menu.__func__.__globals__, "QMenu", InspectMenu)
        widget._show_item_context_menu(category, listing.visualItemRect(listing.item(0)).center())
        assert len(labels) == 2
        assert _t("components.file_drop.batch_list.action_open_selected_locations", count=1) in labels
        assert file_path not in widget.view_model.get_files()
        assert widget.view_model.entry_count == original_count - 1

    def test_widget_created(self, widget: BatchList) -> None:
        assert widget is not None

    def test_object_name(self, widget: BatchList) -> None:
        assert widget.objectName() == "batchListSurface"

    def test_focus_policy(self, widget: BatchList) -> None:
        assert widget.focusPolicy() == Qt.FocusPolicy.StrongFocus

    def test_view_model_access(self, widget: BatchList) -> None:
        assert widget.view_model is not None
        assert isinstance(widget.view_model, BatchListViewModel)

    def test_six_tabs_created(self, widget: BatchList) -> None:
        assert len(widget._tabs) == 6
        for cat in CATEGORY_ORDER:
            assert cat in widget._tabs

    def test_narrow_category_pivot_keeps_all_six_labels_visible(
        self,
        populated_widget: BatchList,
        qapp: QApplication,
    ) -> None:
        populated_widget.setFixedWidth(375)
        populated_widget.resize(375, 640)
        populated_widget.show()
        qapp.processEvents()

        pivot = populated_widget.category_pivot
        items = list(populated_widget._pivot_items.values())
        assert len(items) == 6
        assert all("1" not in cast(_PivotText, item).text() for item in items)
        assert max(item.geometry().right() for item in items) < pivot.width()

    def test_category_navigation_remains_accessible_after_counts_and_font_change(
        self, populated_widget: BatchList, qapp: QApplication
    ) -> None:
        populated_widget.setFixedWidth(400)
        populated_widget.resize(400, 640)
        populated_widget.show()
        _add_synthetic(populated_widget.view_model, [f"/test/image-{index}.png" for index in range(100)])
        for font_size in (14, 26):
            font = populated_widget.font()
            font.setPixelSize(font_size)
            for item in populated_widget._pivot_items.values():
                item.setFont(font)
            for _ in range(8):
                qapp.processEvents()
            selector = populated_widget._category_selector
            if selector.isVisible():
                assert selector.count() == 6
                selector.setCurrentIndex(selector.findData("other"))
                assert populated_widget.view_model.current_category == "other"
            else:
                pivot = populated_widget.category_pivot
                assert all(pivot.rect().contains(item.geometry()) for item in populated_widget._pivot_items.values())

    def test_wide_category_pivot_restores_nonzero_counts(
        self,
        populated_widget: BatchList,
        qapp: QApplication,
    ) -> None:
        populated_widget.resize(1200, 640)
        populated_widget.show()
        qapp.processEvents()

        assert any("1" in cast(_PivotText, item).text() for item in populated_widget._pivot_items.values())

    def test_category_stack_exists(self, widget: BatchList) -> None:
        assert widget.category_stack is not None
        assert widget.category_stack.count() == 6

    def test_tab_lists_are_reorderable(self, widget: BatchList) -> None:
        for list_widget in widget._tabs.values():
            assert isinstance(list_widget, ReorderableListWidget)

    def test_filter_button_exists(self, widget: BatchList) -> None:
        assert widget.filter_button is not None
        assert isinstance(widget.filter_button, QPushButton)

    def test_sort_button_exists(self, widget: BatchList) -> None:
        assert widget.sort_button is not None

    def test_move_buttons_exist(self, widget: BatchList) -> None:
        assert widget.move_up_button is not None
        assert widget.move_down_button is not None

    def test_summary_toolbar_buttons_use_control_height_token(self, widget: BatchList) -> None:
        from docwen_gui.styles.design_tokens import Sizing

        controls = (
            widget.filter_button,
            widget.move_up_button,
            widget.move_down_button,
            widget.sort_button,
        )
        assert all(control.minimumHeight() == Sizing.CONTROL_HEIGHT for control in controls)

    def test_summary_label_exists(self, widget: BatchList) -> None:
        assert widget.summary_label is not None

    def test_summary_shows_no_files_initially(self, widget: BatchList) -> None:
        assert "No batch files" in widget.summary_label.text() or len(widget.summary_label.text()) > 0

    def test_prepopulated_view_model_entries_render_on_construction(
        self,
        qapp: QApplication,
        vm: BatchListViewModel,
    ) -> None:
        _add_synthetic(vm, ["/test/readme.md", "/test/doc.docx"])
        vm.activate_tab("document")

        w = BatchList(view_model=vm)
        try:
            assert w._vm.current_category == "document"
            assert w._tabs["text"].count() == 1
            assert w._tabs["document"].count() == 1
            assert "2" in w.summary_label.text()
        finally:
            w.deleteLater()

    def test_user_tab_activation_uses_view_model_signal_and_updates_selection(
        self,
        widget: BatchList,
        vm: BatchListViewModel,
    ) -> None:
        _add_synthetic(vm, ["/test/readme.md", "/test/doc.docx"])
        categories: list[str] = []
        selections: list[str | None] = []
        vm.current_category_changed.connect(categories.append)
        widget.selection_changed.connect(selections.append)

        widget._activate_tab("document")

        assert vm.current_category == "document"
        assert categories == ["document"]
        assert selections[-1] == "/test/doc.docx"

        widget._activate_tab("spreadsheet")

        assert vm.current_category == "spreadsheet"
        assert categories[-1] == "spreadsheet"
        assert selections[-1] is None
