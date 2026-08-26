"""Focused tests split from test_numbering_editors_and_batch_dialog.py."""

from __future__ import annotations

import pytest

from ._numbering_editors_and_batch_dialog_support import (
    QApplication,
)

pytestmark = pytest.mark.gui


class TestTextTabButtonLabels:
    """Verify the TextTab editor buttons are properly labeled and connected."""

    def test_editor_buttons_have_i18n_labels(self, qapp: QApplication) -> None:
        """Buttons on TextTab use i18n strings, not hardcoded English."""
        from PySide6.QtWidgets import QPushButton

        from docwen_gui.i18n import get_locale, set_locale, t
        from docwen_gui.models.settings_config import SettingsConfig
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.text_tab import TextTab

        previous_locale = get_locale()
        set_locale("zh_CN")
        try:
            vm = SettingsViewModel(config=SettingsConfig())
            tab = TextTab(vm)

            all_labels = {btn.text() for btn in tab.findChildren(QPushButton)}
            add_label = t("settings.text.edit_numbering_add")
            clean_label = t("settings.text.edit_numbering_clean")

            assert add_label == "编辑小标题序号新增方案"
            assert clean_label == "编辑小标题序号清理规则"
            assert add_label in all_labels, f"Add button label '{add_label}' not found in {sorted(all_labels)}"
            assert clean_label in all_labels, f"Clean button label '{clean_label}' not found in {sorted(all_labels)}"
            assert "Edit Numbering Addition Schemes" not in all_labels
            assert "Edit Numbering Removal Rules" not in all_labels
        finally:
            set_locale(previous_locale)

    def test_text_tab_scheme_combo_uses_view_model_ids(self, qapp: QApplication) -> None:
        """TextTab must project its injected ViewModel source, not a global loader."""
        from docwen_gui.models.settings_config import SettingsConfig
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.text_tab import TextTab

        config = SettingsConfig()
        config.text.numbering_schemes = {
            "settings": {"default_scheme": "first", "order": ["first", "second"]},
            "schemes": {
                "first": {"name": "First"},
                "second": {"name": "Second"},
            },
        }
        vm = SettingsViewModel(config=config)
        tab = TextTab(vm)

        combo = tab._scheme_combo
        values = [combo.itemData(i) for i in range(combo.count())]
        assert values == ["first", "second"]

    def test_text_tab_scheme_combo_rebuilds_from_updated_view_model(self, qapp: QApplication) -> None:
        """A reconciled editor source should rebuild the combo from current draft data."""
        from docwen_gui.models.settings_config import SettingsConfig
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.text_tab import TextTab

        config = SettingsConfig()
        config.text.numbering_schemes = {
            "settings": {"default_scheme": "custom_one", "order": ["custom_one", "custom_two"]},
            "schemes": {
                "custom_one": {"name": "Custom One"},
                "custom_two": {"name": "Custom Two"},
            },
        }
        vm = SettingsViewModel(config=config)
        tab = TextTab(vm)

        vm.set_field(
            "text",
            "numbering_schemes",
            {
                "settings": {"default_scheme": "custom_three", "order": ["custom_three"]},
                "schemes": {"custom_three": {"name": "Custom Three"}},
            },
        )
        tab.reload_from_config()

        combo = tab._scheme_combo
        assert [combo.itemData(i) for i in range(combo.count())] == ["custom_three"]
