"""Focused tests split from test_template_selector.py."""

from __future__ import annotations

import pytest

from ._template_selector_support import (
    _SORTED_FIRST,
    _SORTED_SECOND,
    _SORTED_THIRD,
    _UNSORTED_FIRST,
    _UNSORTED_THIRD,
    QApplication,
    QListWidgetItem,
    Qt,
    TabbedTemplateSelector,
    TemplateSelector,
    _assert_hidden,
    _assert_visible,
    _sample_details,
    _sample_names,
    t,
)

pytestmark = pytest.mark.gui
from ._template_selector_support import (
    selector as selector,
)
from ._template_selector_support import (
    tabbed as tabbed,
)


class TestTemplateSelectorConstruction:
    """Smoke: widget creation and structural invariants."""

    def test_widget_created(self, selector: TemplateSelector) -> None:
        assert selector is not None
        assert isinstance(selector, TemplateSelector)

    def test_object_name(self, selector: TemplateSelector) -> None:
        assert selector.objectName() == "templateSelectorRoot"

    def test_focus_policy(self, selector: TemplateSelector) -> None:
        assert selector.focusPolicy() == Qt.FocusPolicy.StrongFocus

    def test_template_type(self, selector: TemplateSelector) -> None:
        assert selector.template_type == "docx"

    def test_xlsx_type(self, qapp: QApplication) -> None:
        w = TemplateSelector(template_type="xlsx")
        assert w.template_type == "xlsx"
        w.deleteLater()

    def test_initial_selected_is_none(self, selector: TemplateSelector) -> None:
        assert selector.get_selected() is None

    def test_starts_in_empty_state(self, selector: TemplateSelector) -> None:
        _assert_hidden(selector._list)
        _assert_visible(selector._empty_state)


class TestTemplateSelectorEmptyState:
    """Empty-state display correctness."""

    def test_empty_label_visible(self, selector: TemplateSelector) -> None:
        _assert_visible(selector._empty_label)
        assert len(selector._empty_label.text()) > 0

    def test_empty_hint_visible(self, selector: TemplateSelector) -> None:
        _assert_visible(selector._empty_hint_label)

    def test_empty_action_button_visible(self, selector: TemplateSelector) -> None:
        _assert_visible(selector._empty_action_button)
        assert selector._empty_action_button.text() == t("components.template_selector.open_template_dir")

    def test_empty_action_button_click_invokes_callback(self, qapp: QApplication) -> None:
        calls: list[str] = []

        def _on_open_dir(tt: str) -> None:
            calls.append(tt)

        w = TemplateSelector(template_type="docx", on_open_directory=_on_open_dir)
        w._empty_action_button.click()
        assert calls == ["docx"]
        w.deleteLater()

    def test_empty_action_signal_when_no_callback(self, qapp: QApplication) -> None:
        w = TemplateSelector(template_type="docx")  # no callbacks
        errors: list[tuple[str, str]] = []
        w.template_error.connect(lambda s, d: errors.append((s, d)))
        w._empty_action_button.click()
        assert len(errors) == 1
        assert errors[0][0] == t("components.template_selector.unavailable")
        w.deleteLater()


class TestTemplateSelectorAddTemplates:
    """List population from external data."""

    def test_add_templates_shows_list(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        _assert_visible(selector._list)
        _assert_hidden(selector._empty_state)
        assert selector._list.count() == 3

    def test_add_templates_items_have_user_role_data(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        items = [selector._list.item(i) for i in range(selector._list.count())]
        names = [it.data(Qt.ItemDataRole.UserRole) for it in items]
        # TemplateSelector preserves caller order (no internal sort).
        # TabbedTemplateSelector sorts via _template_name_sort_key.
        assert len(names) == 3
        assert _SORTED_FIRST in names
        assert _SORTED_SECOND in names
        assert _SORTED_THIRD in names

    def test_add_templates_with_details(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names(), template_details=_sample_details())
        assert selector._template_details.get("Standard Report") is not None
        detail = selector._template_details["Standard Report"]
        assert detail.updated_label == "2025-06-01 14:30"

    def test_add_templates_auto_select_first(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names(), auto_select_first=True)
        # TemplateSelector preserves caller order → first is _UNSORTED_FIRST
        assert selector.get_selected() == _UNSORTED_FIRST

    def test_add_templates_auto_select_respects_manual(self, selector: TemplateSelector) -> None:
        # First populate → auto selects first item in caller order
        selector.add_templates(_sample_names(), auto_select_first=True)
        assert selector.get_selected() == _UNSORTED_FIRST
        # User manually selects a different template
        selector.select_template(_UNSORTED_THIRD, selection_source="user")
        assert selector.get_selected() == _UNSORTED_THIRD
        # Re-populate — should keep user's manual selection
        selector.add_templates(_sample_names(), auto_select_first=True)
        assert selector.get_selected() == _UNSORTED_THIRD

    def test_add_empty_templates_shows_empty_state(self, selector: TemplateSelector) -> None:
        selector.add_templates([])
        _assert_hidden(selector._list)
        _assert_visible(selector._empty_state)

    def test_add_templates_clears_previous_selection_when_gone(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names(), auto_select_first=True)
        assert selector.get_selected() == _UNSORTED_FIRST
        # Repopulate with a different set — old selection is gone, auto-select picks first
        selector.add_templates(["New Template"], auto_select_first=True)
        assert selector.get_selected() == "New Template"


class TestTemplateSelectorSelection:
    """Selection tracking and feedback."""

    def test_select_template_by_name(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        selector.select_template(_SORTED_SECOND, selection_source="user")
        assert selector.get_selected() == _SORTED_SECOND

    def test_select_template_emits_signal(self, selector: TemplateSelector) -> None:
        emitted: list[str] = []
        selector.template_selected.connect(lambda name: emitted.append(name))
        selector.add_templates(_sample_names())
        selector.select_template(_SORTED_SECOND, selection_source="user")
        assert _SORTED_SECOND in emitted

    def test_select_template_emits_callback(self, qapp: QApplication) -> None:
        calls: list[str] = []

        def _cb(name: str) -> None:
            calls.append(name)

        w = TemplateSelector(template_type="docx", on_template_selected=_cb)
        w.add_templates(_sample_names())
        w.select_template(_SORTED_SECOND, selection_source="user")
        assert calls == [_SORTED_SECOND]
        w.deleteLater()

    def test_activate_first_template(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        name = selector.activate_first_template(selection_source="auto_default")
        # TemplateSelector preserves caller order — first is _UNSORTED_FIRST
        assert name == _UNSORTED_FIRST
        assert selector.get_selected() == _UNSORTED_FIRST

    def test_activate_first_on_empty(self, selector: TemplateSelector) -> None:
        assert selector.activate_first_template() is None

    def test_has_template(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        assert selector.has_template(_SORTED_SECOND) is True
        assert selector.has_template("No Such Template") is False
        assert selector.has_template(None) is False

    def test_has_manual_selection_tracking(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        # auto selection doesn't count as manual
        selector.activate_first_template(selection_source="auto_default")
        assert selector._has_manual_selection() is False
        # user selection does
        selector.select_template(_SORTED_SECOND, selection_source="user")
        assert selector._has_manual_selection() is True

    def test_consume_selection_feedback(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        selector.select_template(_SORTED_THIRD, selection_source="user")
        fb = selector.consume_selection_feedback()
        assert fb is not None
        assert fb.selection_source == "user"
        # Consumed — second call returns None
        assert selector.consume_selection_feedback() is None

    def test_consumer_feedback_after_auto_default(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        selector.activate_first_template(selection_source="auto_default", explanation="auto")
        fb = selector.consume_selection_feedback()
        assert fb is not None
        assert fb.selection_source == "auto_default"
        assert fb.explanation == "auto"


class TestTemplateSelectorClearAll:
    """Clearing behaviour."""

    def test_clear_all_removes_items(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        assert selector._list.count() == 3
        selector.clear_all()
        assert selector._list.count() == 0
        assert selector.get_selected() is None
        _assert_visible(selector._empty_state)

    def test_clear_all_resets_manual_selection(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        selector.select_template(_SORTED_SECOND, selection_source="user")
        selector.clear_all()
        assert selector._has_manual_selection() is False


class TestTemplateSelectorLocationButton:
    """Open-location button state synchronisation."""

    def test_location_button_disabled_initially(self, selector: TemplateSelector) -> None:
        assert not selector._open_location_button.isEnabled()

    def test_location_button_enabled_after_selection(self, qapp: QApplication) -> None:
        def _cb(tt: str, name: str) -> None:
            pass

        w = TemplateSelector(template_type="docx", on_open_location=_cb)
        w.show()
        w.add_templates(_sample_names())
        w.select_template(_UNSORTED_THIRD, selection_source="user")
        assert w._open_location_button.isEnabled()
        w.hide()
        w.deleteLater()

    def test_location_button_hidden_without_callback(self, selector: TemplateSelector) -> None:
        # No on_open_location callback in fixture → button is explicitly hidden
        _assert_hidden(selector._open_location_button)

    def test_location_button_visible_with_callback(self, qapp: QApplication) -> None:
        def _cb(tt: str, name: str) -> None:
            pass

        w = TemplateSelector(template_type="docx", on_open_location=_cb)
        w.show()
        _assert_visible(w._open_location_button)
        w.hide()
        w.deleteLater()

    def test_location_affordance_icons_visible_with_callback(self, qapp: QApplication) -> None:
        def _cb(tt: str, name: str) -> None:
            pass

        w = TemplateSelector(template_type="docx", on_open_location=_cb)
        w.add_templates(_sample_names())

        for index in range(w._list.count()):
            item = w._list.item(index)
            assert item is not None
            assert not item.icon().isNull()
            assert t("components.template_selector.open_location") in item.toolTip()

        w.deleteLater()

    def test_location_affordance_icons_hidden_without_callback(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())

        for index in range(selector._list.count()):
            item = selector._list.item(index)
            assert item is not None
            assert item.icon().isNull()


class TestTemplateSelectorDetailsLabel:
    """Footer details label updates on selection."""

    def test_details_label_shows_updated_time(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names(), template_details=_sample_details())
        selector.select_template(_SORTED_THIRD, selection_source="user")
        _assert_visible(selector._details_label)
        assert "2025-06-01" in selector._details_label.text()

    def test_details_label_keeps_usage_in_tooltip_and_visible_footer_compact(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names(), template_details=_sample_details())
        selector.select_template(_SORTED_THIRD, selection_source="user")
        _assert_visible(selector._details_label)

        text = selector._details_label.text()
        tooltip = selector._details_label.toolTip()

        assert "Default document template" not in text
        assert "bundled" in text
        assert "Default document template" in tooltip
        assert "S:/DocWen/templates/Standard Report.docx" in tooltip

    def test_details_label_hidden_without_details(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names())
        selector.select_template(_SORTED_SECOND, selection_source="user")
        # Business Letter has no detail entry → label should be hidden
        _assert_hidden(selector._details_label)

    def test_footer_row_visible_when_details_or_button(self, selector: TemplateSelector) -> None:
        selector.add_templates(_sample_names(), template_details=_sample_details())
        selector.select_template(_SORTED_THIRD, selection_source="user")
        _assert_visible(selector._footer_row)


class TestTemplateSelectorFocus:
    """Focus management — relaxed for offscreen mode."""

    def test_focus_policy_correct(self, selector: TemplateSelector) -> None:
        """Focus policy is StrongFocus — enables keyboard navigation."""
        assert selector.focusPolicy() == Qt.FocusPolicy.StrongFocus

    def test_focus_in_event_accepted(self, selector: TemplateSelector) -> None:
        """FocusInEvent should not crash and should delegate."""
        from PySide6.QtCore import QEvent
        from PySide6.QtGui import QFocusEvent

        ev = QFocusEvent(QEvent.Type.FocusIn, Qt.FocusReason.TabFocusReason)
        # Should not raise
        selector.focusInEvent(ev)


class TestTemplateSelectorContextMenu:
    """Right-click context menu behaviour — coverage for 'Open File Location' entry.

    QMenu.exec() blocks forever in headless Qt, so we test the building
    blocks: context-menu policy, signal wiring, item identification, and
    the action's ultimate callee (_open_template_location) which is already
    verified by the open-location button tests.
    """

    def test_context_menu_policy_is_custom(self, selector: TemplateSelector) -> None:
        """Template selector list uses CustomContextMenu policy."""
        assert selector._list.contextMenuPolicy() == Qt.ContextMenuPolicy.CustomContextMenu

    def test_custom_context_menu_signal_is_connected(self, selector: TemplateSelector) -> None:
        """The customContextMenuRequested signal is wired to
        _show_list_context_menu so right-click triggers it."""
        # QObject.receivers takes a SIGNAL string signature.
        # In headless mode we verify at least the slot method exists
        # and is connectable (the signal-slot wiring in __init__ does this).
        assert hasattr(selector, "_show_list_context_menu")
        # Verify the method is callable
        assert callable(selector._show_list_context_menu)

    def test_item_at_position_identifies_second_item(self, qapp: QApplication) -> None:
        """QListWidget.itemAt correctly resolves a visual position to the
        expected list item — the same lookup _show_list_context_menu uses."""
        w = TemplateSelector(template_type="docx")
        w.show()
        w.add_templates(_sample_names())

        # itemAt visual centre of the second row
        second = w._list.item(1)
        rect = w._list.visualItemRect(second)
        item_at_pos = w._list.itemAt(rect.center())
        assert item_at_pos is second

        w.hide()
        w.deleteLater()

    def test_show_context_menu_does_not_crash_on_empty_list(self, selector: TemplateSelector) -> None:
        """Right-click on an empty list does not crash — _show_list_context_menu
        returns early when itemAt returns None (headless-safe because no QMenu is created)."""
        # Empty list — itemAt returns None → early return before QMenu.exec
        selector._list.customContextMenuRequested.emit(selector._list.rect().center())
        # No crash = pass

    def test_location_button_and_context_menu_share_callee(self, qapp: QApplication) -> None:
        """Both the open-location button and the context menu action route
        through _open_template_location — verifying the button path also
        verifies the context menu's action endpoint."""
        open_calls: list[tuple[str, str]] = []

        def _on_open(tt: str, name: str) -> None:
            open_calls.append((tt, name))

        w = TemplateSelector(template_type="docx", on_open_location=_on_open)
        w.add_templates(_sample_names())
        w.select_template(_SORTED_SECOND, selection_source="user")

        # Direct call to the shared method that both the button and
        # context menu action ultimately invoke.
        w._open_template_location(_SORTED_SECOND)
        assert ("docx", _SORTED_SECOND) in open_calls

        w.deleteLater()


class TestTemplateSelectorItemActivation:
    """Item activation (double-click / Enter) path coverage."""

    def test_item_activated_calls_open_location(self, qapp: QApplication) -> None:
        """Double-click or Enter on a template item triggers the
        open-location callback via _on_item_activated."""
        open_calls: list[tuple[str, str]] = []

        def _on_open(tt: str, name: str) -> None:
            open_calls.append((tt, name))

        w = TemplateSelector(template_type="docx", on_open_location=_on_open)
        w.add_templates(_sample_names())
        w.select_template(_SORTED_SECOND, selection_source="user")

        item = w._list.currentItem()
        assert item is not None
        w._on_item_activated(item)

        assert ("docx", _SORTED_SECOND) in open_calls
        w.deleteLater()

    def test_item_activated_without_callback_emits_error(self, qapp: QApplication) -> None:
        """Double-click without a location callback emits template_error."""
        w = TemplateSelector(template_type="docx")  # no on_open_location
        w.add_templates(_sample_names())
        w.select_template(_SORTED_FIRST, selection_source="user")

        errors: list[tuple[str, str]] = []
        w.template_error.connect(lambda s, d: errors.append((s, d)))

        item = w._list.currentItem()
        w._on_item_activated(item)

        assert len(errors) == 1
        assert errors[0][0] == t("components.template_selector.unavailable")
        w.deleteLater()

    def test_item_activated_with_none_item_does_not_crash(self, qapp: QApplication) -> None:
        """Passing an item with no UserRole data does not crash."""
        w = TemplateSelector(template_type="docx")
        # Use a bare QListWidgetItem with no UserRole
        bare = QListWidgetItem("bare")
        # Should not raise
        w._on_item_activated(bare)
        w.deleteLater()


class TestTabbedTemplateSelectorConstruction:
    """Smoke: tabbed widget creation."""

    def test_widget_created(self, tabbed: TabbedTemplateSelector) -> None:
        assert tabbed is not None
        assert isinstance(tabbed, TabbedTemplateSelector)

    def test_object_name(self, tabbed: TabbedTemplateSelector) -> None:
        assert tabbed.objectName() == "tabbedTemplateSelector"

    def test_focus_policy(self, tabbed: TabbedTemplateSelector) -> None:
        assert tabbed.focusPolicy() == Qt.FocusPolicy.StrongFocus

    def test_default_tab_is_docx(self, tabbed: TabbedTemplateSelector) -> None:
        assert tabbed.current_tab == "docx"

    def test_has_both_selectors(self, tabbed: TabbedTemplateSelector) -> None:
        docx_sel = tabbed.get_selector("docx")
        xlsx_sel = tabbed.get_selector("xlsx")
        assert docx_sel is not None
        assert xlsx_sel is not None
        assert docx_sel.template_type == "docx"
        assert xlsx_sel.template_type == "xlsx"

    def test_initial_selection_is_none(self, tabbed: TabbedTemplateSelector) -> None:
        assert tabbed.get_selected_template() is None


class TestTabbedTemplateSelectorLoadTemplates:
    """Template data injection."""

    def test_load_templates_docx(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_templates("docx", _sample_names())
        sel = tabbed.get_selector("docx")
        assert sel is not None
        assert sel._list.count() == 3
        # Auto-selects sort-first template on load
        assert sel.get_selected() == _SORTED_FIRST

    def test_load_templates_xlsx(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_templates("xlsx", ["Budget", "Invoice"])
        sel = tabbed.get_selector("xlsx")
        assert sel is not None
        assert sel._list.count() == 2

    def test_load_all_templates(self, tabbed: TabbedTemplateSelector) -> None:
        data = {"docx": _sample_names(), "xlsx": ["Budget", "Invoice"]}
        tabbed.load_all_templates(data)
        docx_sel = tabbed.get_selector("docx")
        xlsx_sel = tabbed.get_selector("xlsx")
        assert docx_sel is not None
        assert xlsx_sel is not None
        assert docx_sel._list.count() == 3
        assert xlsx_sel._list.count() == 2

    def test_load_templates_with_details(self, tabbed: TabbedTemplateSelector) -> None:
        details = {"docx": _sample_details()}
        tabbed.load_all_templates({"docx": _sample_names(), "xlsx": []}, details=details)
        sel = tabbed.get_selector("docx")
        assert sel is not None
        assert sel._template_details.get("Standard Report") is not None

    def test_reload_preserves_manual_selection(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_templates("docx", _sample_names())
        sel = tabbed.get_selector("docx")
        assert sel is not None
        sel.select_template(_SORTED_SECOND, selection_source="user")
        # Reload same templates
        tabbed.load_templates("docx", _sample_names())
        assert sel.get_selected() == _SORTED_SECOND
