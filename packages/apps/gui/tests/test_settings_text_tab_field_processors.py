from __future__ import annotations

import pytest
from PySide6.QtWidgets import QCheckBox, QWidget

from docwen_gui.models.settings_config import SettingsConfig, TextConfig
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.text_tab import TextTab

pytestmark = pytest.mark.gui


def _field_processor_config(enabled: bool = True) -> dict:
    return {
        "settings": {"order": ["gongwen"]},
        "processors": {
            "gongwen": {
                "module": "docwen_plugin_markdown.field_processors.gongwen",
                "name": "公文字段优化",
                "description": "附件说明、抄送机关和日期格式化",
                "locales": ["zh_CN"],
                "enabled": enabled,
                "is_system": True,
            }
        },
    }


def test_text_tab_shows_field_processor_card_and_toggle(qapp) -> None:
    config = SettingsConfig(text=TextConfig(field_processors=_field_processor_config(enabled=True)))
    vm = SettingsViewModel(config=config)

    tab = TextTab(vm)

    assert tab.findChild(QWidget, "textFieldProcessorsCard") is not None
    checkbox = tab._field_processor_checkboxes["gongwen"]
    assert isinstance(checkbox, QCheckBox)
    assert checkbox.isChecked() is True

    checkbox.setChecked(False)

    processors = vm.config.text.field_processors["processors"]
    assert processors["gongwen"]["enabled"] is False
    assert vm.is_dirty is True
    tab.close()


def test_text_tab_reload_restores_field_processor_toggle(qapp) -> None:
    config = SettingsConfig(text=TextConfig(field_processors=_field_processor_config(enabled=True)))
    vm = SettingsViewModel(config=config)
    tab = TextTab(vm)

    tab._field_processor_checkboxes["gongwen"].setChecked(False)
    vm.cancel()
    tab.reload_from_config()

    assert tab._field_processor_checkboxes["gongwen"].isChecked() is True
    tab.close()


def test_text_tab_treats_mutable_name_key_as_data_not_catalog_authority(qapp) -> None:
    field_processors = _field_processor_config(enabled=True)
    field_processors["processors"]["gongwen"]["name"] = "Explicit plugin label"
    field_processors["processors"]["gongwen"]["name_key"] = "common.error"
    config = SettingsConfig(text=TextConfig(field_processors=field_processors))
    vm = SettingsViewModel(config=config)

    tab = TextTab(vm)

    assert tab._field_processor_checkboxes["gongwen"].text() == "Explicit plugin label"
    tab.close()
