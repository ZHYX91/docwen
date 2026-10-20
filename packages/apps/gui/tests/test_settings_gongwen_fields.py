from __future__ import annotations

import pytest
from PySide6.QtWidgets import QAbstractButton, QWidget

from docwen_gui.models.settings_config import SettingsConfig, TextConfig
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.text_tab import TextTab

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("language", ["zh_CN", "en_US", "ar_SA"])
def test_gongwen_fields_toggle_has_help_and_is_independent_of_ui_language(qapp, language) -> None:
    config = SettingsConfig(text=TextConfig(gongwen_fields_enabled=True))
    config.gui.language = language
    vm = SettingsViewModel(config=config)
    tab = TextTab(vm)

    card = tab.findChild(QWidget, "textGongwenFieldsCard")
    assert card is not None
    assert tab._gongwen_fields.isChecked()
    assert any(button is not tab._gongwen_fields for button in card.findChildren(QAbstractButton))
    assert tab._gongwen_fields.toolTip()

    tab._gongwen_fields.setChecked(False)
    assert vm.config.text.gongwen_fields_enabled is False
    assert vm.is_dirty
    tab.close()


def test_cancel_restores_gongwen_fields_toggle(qapp) -> None:
    vm = SettingsViewModel(config=SettingsConfig(text=TextConfig(gongwen_fields_enabled=True)))
    tab = TextTab(vm)
    tab._gongwen_fields.setChecked(False)
    vm.cancel()
    tab.reload_from_config()

    assert tab._gongwen_fields.isChecked()
    assert vm.config.text.gongwen_fields_enabled is True
    tab.close()
