from __future__ import annotations

import pytest
from PySide6.QtWidgets import QToolButton, QWidget

from docwen_gui.i18n import get_locale, set_locale, t
from docwen_gui.models.settings_config import SettingsConfig, TextConfig
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.gongwen_tab import GongwenTab

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("language", ["zh_CN", "en_US", "ar_SA"])
def test_gongwen_fields_toggle_has_help_and_is_independent_of_ui_language(qapp, language) -> None:
    previous_locale = get_locale()
    set_locale(language)
    config = SettingsConfig(text=TextConfig(gongwen_fields_enabled=True))
    config.gui.language = language
    vm = SettingsViewModel(config=config)
    tab = GongwenTab(vm)
    try:
        card = tab.findChild(QWidget, "gongwenFromMarkdownCard")
        assert card is not None
        assert tab._gongwen_fields.isChecked()
        info = card.findChild(QToolButton, "settingsInfoButton")
        assert info is not None
        assert info.accessibleName() == t("settings.gongwen.fields_label")
        assert info.toolTip() == tab._gongwen_fields.toolTip() == t("settings.gongwen.fields_help")
        tab._gongwen_fields.setChecked(False)
        assert vm.config.text.gongwen_fields_enabled is False
        assert vm.is_dirty
    finally:
        tab.close()
        set_locale(previous_locale)


def test_cancel_restores_gongwen_fields_toggle(qapp) -> None:
    vm = SettingsViewModel(config=SettingsConfig(text=TextConfig(gongwen_fields_enabled=True)))
    tab = GongwenTab(vm)
    tab._gongwen_fields.setChecked(False)
    vm.cancel()

    assert tab._gongwen_fields.isChecked()
    assert vm.config.text.gongwen_fields_enabled is True
    tab.close()
