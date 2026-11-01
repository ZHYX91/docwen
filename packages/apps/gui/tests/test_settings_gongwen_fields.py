from __future__ import annotations

import pytest
from PySide6.QtWidgets import QToolButton, QWidget

from docwen_gui.i18n import get_locale, set_locale, t
from docwen_gui.models.settings_config import SettingsConfig, TextConfig
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.gongwen_tab import GongwenTab

pytestmark = pytest.mark.gui


def test_numbering_text_and_review_default_on_with_independent_fields_and_linked_controls(qapp) -> None:
    vm = SettingsViewModel(config=SettingsConfig(text=TextConfig(gongwen_fields_enabled=False, add_numbering=False)))
    tab = GongwenTab(vm)
    try:
        assert tab._numbering_text.isChecked()
        assert tab._numbering_comments.isChecked()
        assert tab._numbering_comments.isEnabled()
        assert tab._numbering_comments.toolTip() == t("settings.gongwen.numbering_comments_help")
        tab._numbering_text.setChecked(False)
        assert not tab._numbering_comments.isChecked()
        assert not tab._numbering_comments.isEnabled()
        tab._numbering_text.setChecked(True)
        assert not tab._numbering_comments.isChecked()
        assert tab._numbering_comments.isEnabled()
        tab._numbering_comments.setChecked(True)
        assert vm.config.text.gongwen_numbering_comments is True
        assert vm.config.text.add_numbering is False
        assert vm.config.text.gongwen_fields_enabled is False
        vm.cancel()
        assert tab._numbering_text.isChecked()
        assert tab._numbering_comments.isChecked()
    finally:
        tab.close()


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
        assert tab._attachment_alignment.currentData() == "tab"
        assert tab._attachment_alignment.isEnabled()
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


@pytest.mark.parametrize(
    "language",
    ["zh_CN", "en_US", "ar_SA", "de_DE", "es_ES", "fr_FR", "ja_JP", "ko_KR", "pt_BR", "ru_RU", "vi_VN", "zh_TW"],
)
def test_attachment_alignment_is_a_translated_dependent_choice(qapp, language) -> None:
    previous_locale = get_locale()
    set_locale(language)
    vm = SettingsViewModel(config=SettingsConfig(text=TextConfig(attachment_alignment="spaces")))
    tab = GongwenTab(vm)
    try:
        combo = tab._attachment_alignment
        assert combo.currentData() == "spaces"
        assert combo.accessibleName() == t("settings.gongwen.attachment_alignment_label")
        assert "settings.gongwen." not in combo.accessibleName()
        assert combo.count() == 2
        assert all("settings.gongwen." not in combo.itemText(i) for i in range(combo.count()))
        assert combo.isEnabled()

        combo.setCurrentIndex(combo.findData("tab"))
        assert vm.config.text.attachment_alignment == "tab"
        tab._gongwen_fields.setChecked(False)
        assert not combo.isEnabled()
        assert combo.currentData() == vm.config.text.attachment_alignment == "tab"
        tab._gongwen_fields.setChecked(True)
        assert combo.isEnabled()
        assert combo.currentData() == "tab"

        vm.cancel()
        assert combo.currentData() == vm.config.text.attachment_alignment == "spaces"
        assert not vm.is_dirty
    finally:
        tab.close()
        set_locale(previous_locale)
