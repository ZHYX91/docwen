"""Focused tests split from test_template_selector.py."""

from __future__ import annotations

from ._template_selector_support import (
    _SORTED_SECOND,
    _SORTED_THIRD,
    Any,
    QApplication,
    Qt,
    TabbedTemplateSelector,
    TemplateSelectionFeedback,
    _assert_visible,
    _sample_names,
    pytest,
)

pytestmark = pytest.mark.gui
from ._template_selector_support import (
    selector as selector,
)
from ._template_selector_support import (
    tabbed as tabbed,
)


class TestTabbedTemplateSelectorTabSwitching:
    """Tab switching behaviour."""

    def test_switch_tab_via_method(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed._set_current_tab("xlsx", emit_signal=False)
        assert tabbed.current_tab == "xlsx"

    def test_switch_tab_emits_signal(self, tabbed: TabbedTemplateSelector) -> None:
        emitted: list[tuple[str, str]] = []
        tabbed.tab_changed.connect(lambda new, old: emitted.append((new, old)))
        tabbed._set_current_tab("xlsx", emit_signal=True)
        assert len(emitted) == 1
        assert emitted[0] == ("xlsx", "docx")

    def test_switch_tab_emits_callback(self, qapp: QApplication) -> None:
        calls: list[tuple[str, str]] = []

        def _cb(new: str, old: str) -> None:
            calls.append((new, old))

        w = TabbedTemplateSelector(on_tab_changed=_cb)
        w._set_current_tab("xlsx", emit_signal=True)
        assert calls == [("xlsx", "docx")]
        w.deleteLater()

    def test_switch_tab_preserves_selection(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_all_templates({"docx": _sample_names(), "xlsx": ["Budget", "Invoice"]})
        docx_sel = tabbed.get_selector("docx")
        xlsx_sel = tabbed.get_selector("xlsx")
        assert docx_sel is not None
        assert xlsx_sel is not None

        docx_sel.select_template(_SORTED_SECOND, selection_source="user")
        tabbed._set_current_tab("xlsx", emit_signal=True)
        xlsx_sel.select_template("Invoice", selection_source="user")

        tabbed._set_current_tab("docx", emit_signal=True)
        assert docx_sel.get_selected() == _SORTED_SECOND
        tabbed._set_current_tab("xlsx", emit_signal=True)
        assert xlsx_sel.get_selected() == "Invoice"


class TestTabbedTemplateSelectorGetSelected:
    """Cross-tab selected template retrieval."""

    def test_returns_current_tab_selection(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_templates("docx", _sample_names())
        sel = tabbed.get_selector("docx")
        assert sel is not None
        sel.select_template(_SORTED_SECOND, selection_source="user")
        result = tabbed.get_selected_template()
        assert result == ("docx", _SORTED_SECOND)

    def test_fallback_cross_tab(self, tabbed: TabbedTemplateSelector) -> None:
        # Only xlsx has templates
        tabbed.load_templates("xlsx", ["Budget"])
        xlsx_sel = tabbed.get_selector("xlsx")
        assert xlsx_sel is not None
        xlsx_sel.select_template("Budget", selection_source="user")
        result = tabbed.get_selected_template()
        assert result == ("xlsx", "Budget")


class TestTabbedTemplateSelectorConsumeFeedback:
    """Selection feedback consumption."""

    def test_feedback_on_user_selection(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_templates("docx", _sample_names())
        docx_sel = tabbed.get_selector("docx")
        assert docx_sel is not None
        docx_sel.select_template(_SORTED_SECOND, selection_source="user")
        fb = tabbed.consume_last_selection_feedback()
        assert fb is not None
        tt, name, feedback = fb
        assert tt == "docx"
        assert name == _SORTED_SECOND
        assert feedback.selection_source == "user"

    def test_feedback_consumed_once(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_templates("docx", _sample_names())
        docx_sel = tabbed.get_selector("docx")
        assert docx_sel is not None
        docx_sel.select_template(_SORTED_THIRD, selection_source="user")
        assert tabbed.consume_last_selection_feedback() is not None
        assert tabbed.consume_last_selection_feedback() is None


class TestTabbedTemplateSelectorActivate:
    """Activate-and-select helpers."""

    def test_activate_and_select_populated(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_templates("xlsx", ["Budget", "Invoice"])
        assert tabbed.activate_and_select("xlsx") is True
        assert tabbed.current_tab == "xlsx"
        xlsx_sel = tabbed.get_selector("xlsx")
        assert xlsx_sel is not None
        assert xlsx_sel.get_selected() == "Budget"

    def test_activate_and_select_empty(self, tabbed: TabbedTemplateSelector) -> None:
        assert tabbed.activate_and_select("xlsx") is False

    def test_ensure_preferred_selection_user_first(self, tabbed: TabbedTemplateSelector) -> None:
        tabbed.load_templates("docx", _sample_names())
        docx_sel = tabbed.get_selector("docx")
        assert docx_sel is not None
        docx_sel.select_template(_SORTED_SECOND, selection_source="user")
        # Switch away and back — ensure_preferred should restore manual selection
        tabbed._set_current_tab("xlsx", emit_signal=False)
        assert tabbed.ensure_preferred_selection("docx") is True
        assert tabbed.current_tab == "docx"
        assert docx_sel.get_selected() == _SORTED_SECOND


class TestTabbedTemplateSelectorFocus:
    """Focus delegation — relaxed for offscreen mode."""

    def test_focus_in_event_accepted(self, tabbed: TabbedTemplateSelector) -> None:
        """FocusInEvent should not crash."""
        from PySide6.QtCore import QEvent
        from PySide6.QtGui import QFocusEvent

        tabbed.load_templates("docx", _sample_names())
        ev = QFocusEvent(QEvent.Type.FocusIn, Qt.FocusReason.TabFocusReason)
        # Should not raise
        tabbed.focusInEvent(ev)


class TestUserPathTemplateSelection:
    """End-to-end user path: open → load templates → select → confirm."""

    def test_full_selection_flow(self, qapp: QApplication) -> None:
        """User opens template selector, templates are loaded,
        user clicks to select one, feedback is consumed."""
        selected: list[tuple[str, str]] = []

        def _on_select(tt: str, name: str) -> None:
            selected.append((tt, name))

        tabbed = TabbedTemplateSelector(on_template_selected=_on_select)
        # Simulate external template loader providing data
        docx_names = _sample_names()
        xlsx_names = ["Quarterly Budget", "Invoice Template"]
        tabbed.load_all_templates({"docx": docx_names, "xlsx": xlsx_names})

        # User manually selects a docx template
        docx_sel = tabbed.get_selector("docx")
        assert docx_sel is not None
        docx_sel.select_template(_SORTED_SECOND, selection_source="user")
        assert len(selected) >= 1
        fb = tabbed.consume_last_selection_feedback()
        assert fb is not None
        assert fb[1] == _SORTED_SECOND

        # User manually selects an xlsx template (directly, without tab switch
        # triggering ensure_preferred_selection which would restore docx)
        xlsx_sel = tabbed.get_selector("xlsx")
        assert xlsx_sel is not None
        xlsx_sel.select_template("Invoice Template", selection_source="user")
        # Manual selection is now ("xlsx", "Invoice Template")
        assert tabbed._manual_selection == ("xlsx", "Invoice Template")

        # Switch to xlsx — manual selection is restored
        tabbed._set_current_tab("xlsx", emit_signal=True)
        result = tabbed.get_selected_template()
        assert result is not None
        assert result[1] == "Invoice Template"

        tabbed.deleteLater()

    def test_callback_wiring_user_path(self, qapp: QApplication) -> None:
        """Simulate: external code wires callbacks, loads templates,
        user selects, application receives the callback."""
        selected_calls: list[tuple[str, str]] = []
        open_location_calls: list[tuple[str, str]] = []

        def _on_template_selected(tt: str, name: str) -> None:
            selected_calls.append((tt, name))

        def _on_open_location(tt: str, name: str) -> None:
            open_location_calls.append((tt, name))

        tabbed = TabbedTemplateSelector(
            on_template_selected=_on_template_selected,
            on_open_location=_on_open_location,
        )
        tabbed.show()
        tabbed.load_templates("docx", _sample_names())

        # User selects a template
        docx_sel = tabbed.get_selector("docx")
        assert docx_sel is not None
        docx_sel.select_template(_SORTED_SECOND, selection_source="user")
        assert ("docx", _SORTED_SECOND) in selected_calls

        # Verify the selector's open location button is visible and enabled
        _assert_visible(docx_sel._open_location_button)
        assert docx_sel._open_location_button.isEnabled()

        docx_sel._open_location_button.click()
        assert ("docx", _SORTED_SECOND) in open_location_calls

        tabbed.hide()
        tabbed.deleteLater()

    def test_selection_callback_can_be_replaced_and_disabled(
        self,
        tabbed: TabbedTemplateSelector,
    ) -> None:
        callback_a: list[tuple[str, str]] = []
        callback_b: list[tuple[str, str]] = []
        signal_calls: list[tuple[str, str]] = []
        tabbed.load_templates("docx", _sample_names())
        tabbed.template_selected.connect(lambda kind, name: signal_calls.append((kind, name)))

        tabbed.set_selection_callback(lambda kind, name: callback_a.append((kind, name)))
        tabbed.set_selection_callback(lambda kind, name: callback_b.append((kind, name)))
        selector = tabbed.get_selector("docx")
        assert selector is not None
        selector.select_template(_SORTED_SECOND, selection_source="user")

        assert callback_a == []
        assert callback_b == [("docx", _SORTED_SECOND)]
        assert signal_calls == [("docx", _SORTED_SECOND)]

        tabbed.set_selection_callback(None)
        selector.select_template(_SORTED_THIRD, selection_source="user")

        assert callback_b == [("docx", _SORTED_SECOND)]
        assert signal_calls == [
            ("docx", _SORTED_SECOND),
            ("docx", _SORTED_THIRD),
        ]


def test_template_selector_importable() -> None:
    from docwen_gui.widgets.template_selector import TemplateSelector as TS

    assert TS is not None


def test_tabbed_template_selector_importable() -> None:
    from docwen_gui.widgets.template_selector_tabbed import (
        TabbedTemplateSelector as TTS,
    )

    assert TTS is not None


def test_data_types_importable() -> None:
    from docwen_gui.widgets.template_selector import (
        TemplateItemDetails,
    )

    assert TemplateItemDetails(usage_hint="test") is not None
    assert TemplateSelectionFeedback(selection_source="user") is not None


class TestSettingsEntryTemplateSelection:
    """Integration tests that prove the settings page user path is wired.

    These tests validate:
    - TextTab instantiates a real TabbedTemplateSelector (not a hardcoded QComboBox).
    - Template data flows from SettingsViewModel to the widget.
    - User template selection flows back to the ViewModel.
    - The old _template_combo is no longer the active template UI.
    """

    @pytest.fixture
    def settings_vm(self) -> Any:
        from docwen_gui.view_models.settings_vm import SettingsViewModel

        return SettingsViewModel()

    @pytest.fixture
    def sample_template_data(self) -> dict[str, list[str]]:
        return {
            "docx": ["Standard Report", "Academic Paper", "Business Letter"],
            "xlsx": ["Quarterly Budget", "Invoice Template"],
        }

    def test_text_tab_creates_tabbed_template_selector(self, qapp: QApplication, settings_vm: Any) -> None:
        """TextTab instantiates a real TabbedTemplateSelector, not a dead QComboBox."""
        from docwen_gui.widgets.settings.text_tab import TextTab
        from docwen_gui.widgets.template_selector_tabbed import TabbedTemplateSelector as TTS

        tab = TextTab(settings_vm)
        assert tab._template_selector is not None
        assert isinstance(tab._template_selector, TTS)
        tab.deleteLater()

    def test_text_tab_old_combo_is_inactive(self, qapp: QApplication, settings_vm: Any) -> None:
        """The old hardcoded _template_combo is left as None — replaced by the selector."""
        from docwen_gui.widgets.settings.text_tab import TextTab

        tab = TextTab(settings_vm)
        assert tab._template_combo is None
        assert tab._template_selector is not None
        tab.deleteLater()

    def test_vm_template_signal_exists(self, settings_vm: Any) -> None:
        """SettingsViewModel exposes template_lists_changed and
        template_selection_changed signals."""
        from PySide6.QtCore import Signal

        assert isinstance(settings_vm.template_lists_changed, Signal)
        assert isinstance(settings_vm.template_selection_changed, Signal)

    def test_vm_set_templates_emits_signal(self, settings_vm: Any) -> None:
        """Calling set_templates() emits template_lists_changed."""
        emitted: list[object] = []
        settings_vm.template_lists_changed.connect(lambda d: emitted.append(d))
        data = {"docx": ["A", "B"], "xlsx": ["C"]}
        settings_vm.set_templates(data)
        assert len(emitted) == 1
        assert emitted[0] == data

    def test_vm_get_templates_roundtrip(self, settings_vm: Any) -> None:
        """get_templates() returns what was set via set_templates()."""
        data = {"docx": ["A", "B"], "xlsx": ["C"]}
        settings_vm.set_templates(data)
        result = settings_vm.get_templates()
        assert result == data

    def test_vm_select_template_emits_signal(self, settings_vm: Any) -> None:
        """select_template() emits template_selection_changed and stores selection."""
        emitted: list[tuple[str, str]] = []
        settings_vm.template_selection_changed.connect(lambda tt, name: emitted.append((tt, name)))
        settings_vm.select_template("docx", "Standard Report")
        assert len(emitted) == 1
        assert emitted[0] == ("docx", "Standard Report")
        assert settings_vm.selected_templates == {"docx": "Standard Report"}

    def test_template_data_flows_from_vm_to_text_tab_selector(
        self,
        qapp: QApplication,
        settings_vm: Any,
        sample_template_data: dict[str, list[str]],
    ) -> None:
        """When VM.set_templates() is called, TextTab's TabbedTemplateSelector
        receives the data via the template_lists_changed signal."""
        from docwen_gui.widgets.settings.text_tab import TextTab

        # Create tab first (connects to VM signals)
        tab = TextTab(settings_vm)
        # Inject templates via VM — signal fires, tab receives
        settings_vm.set_templates(sample_template_data)
        qapp.processEvents()

        selector = tab._template_selector
        assert selector is not None
        # Verify docx templates are loaded
        docx_sel = selector.get_selector("docx")
        assert docx_sel is not None
        assert docx_sel._list.count() == 3
        # Verify xlsx templates are loaded
        xlsx_sel = selector.get_selector("xlsx")
        assert xlsx_sel is not None
        assert xlsx_sel._list.count() == 2

        tab.deleteLater()

    def test_user_template_selection_flows_to_vm(
        self,
        qapp: QApplication,
        settings_vm: Any,
        sample_template_data: dict[str, list[str]],
    ) -> None:
        """When user selects a template in the widget, the ViewModel
        receives the selection and the config field is updated."""
        from docwen_gui.widgets.settings.text_tab import TextTab

        tab = TextTab(settings_vm)
        settings_vm.set_templates(sample_template_data)
        qapp.processEvents()

        # User selects a docx template
        selector = tab._template_selector
        docx_sel = selector.get_selector("docx")
        assert docx_sel is not None
        docx_sel.select_template("Academic Paper", selection_source="user")

        # Verify VM received the selection
        assert settings_vm.selected_templates.get("docx") == "Academic Paper"
        # Verify config field was updated
        assert settings_vm.get_field("gui", "md_default_template") == "docx"

        tab.deleteLater()

    def test_full_settings_entry_user_path(
        self,
        qapp: QApplication,
        settings_vm: Any,
        sample_template_data: dict[str, list[str]],
    ) -> None:
        """End-to-end user path: open settings → templates load → user selects →
        selection is persisted in ViewModel."""
        from docwen_gui.widgets.settings.text_tab import TextTab

        # 1. Application opens settings dialog, creates TextTab with VM
        tab = TextTab(settings_vm)

        # 2. Application injects template data (e.g. from TemplateRegistry)
        settings_vm.set_templates(sample_template_data)
        qapp.processEvents()

        # 3. Verify templates loaded correctly in both tabs
        selector = tab._template_selector
        docx_sel = selector.get_selector("docx")
        xlsx_sel = selector.get_selector("xlsx")
        assert docx_sel is not None
        assert xlsx_sel is not None
        assert docx_sel._list.count() == 3
        assert xlsx_sel._list.count() == 2
        # TabbedTemplateSelector auto-selects first sorted template
        assert selector.get_selected_template() is not None

        # 4. User switches to xlsx tab and selects a template
        selector._set_current_tab("xlsx", emit_signal=True)
        qapp.processEvents()
        xlsx_sel.select_template("Invoice Template", selection_source="user")
        assert settings_vm.selected_templates.get("xlsx") == "Invoice Template"

        # 5. verify feedback consumption
        fb = selector.consume_last_selection_feedback()
        assert fb is not None
        assert fb[0] == "xlsx"
        assert fb[1] == "Invoice Template"
        assert fb[2].selection_source == "user"

        tab.deleteLater()
