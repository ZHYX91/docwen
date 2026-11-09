"""General heading-number review controls remain reachable in every UI language."""

from __future__ import annotations

import pytest
from PySide6.QtWidgets import QCheckBox, QLabel, QWidget

from docwen_gui.i18n import get_locale, set_locale, t
from docwen_gui.models.settings_config import SettingsConfig, TextConfig
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.text_tab import TextTab

pytestmark = pytest.mark.gui


@pytest.mark.parametrize(
    "language",
    ["zh_CN", "en_US", "zh_TW", "ar_SA", "de_DE", "es_ES", "fr_FR", "ja_JP", "ko_KR", "pt_BR", "ru_RU", "vi_VN"],
)
def test_numbering_review_is_general_and_preserves_saved_mode_and_cancel(qapp, language: str) -> None:
    previous_locale = get_locale()
    set_locale(language)
    config = SettingsConfig(
        text=TextConfig(gongwen_fields_enabled=False, add_numbering=False, heading_numbering_render_mode="word_native")
    )
    config.gui.language = language
    vm = SettingsViewModel(config=config)
    tab = TextTab(vm)
    try:
        card = tab.findChild(QWidget, "textNumberingCard")
        assert card is not None
        assert tab.findChild(QCheckBox, "textNumberingText") is None
        assert card.isAncestorOf(tab._numbering_comments)
        assert not tab._numbering_comments.isChecked()
        assert tab._numbering_comments.isEnabled()
        assert tab._numbering_comments.toolTip() == t("settings.text.numbering_comments_help")
        assert tab._render_mode_combo.toolTip() == t("settings.text.render_mode_hint")
        assert not any(label.text() == t("settings.text.render_mode_hint") for label in card.findChildren(QLabel))

        assert vm.config.text.gongwen_numbering_comments is False
        tab._add_numbering.setChecked(True)
        tab.set_combo_data(tab._render_mode_combo, "text")
        assert vm.config.text.heading_numbering_render_mode == "text"
        assert not tab._numbering_comments.isChecked()
        tab.set_combo_data(tab._render_mode_combo, "word_native")
        tab._numbering_comments.setChecked(True)
        assert vm.config.text.gongwen_numbering_comments is True
        assert vm.config.text.gongwen_fields_enabled is False
        assert vm.config.text.heading_numbering_render_mode == "word_native"

        vm.cancel()
        assert not tab._numbering_comments.isChecked()
        assert tab._render_mode_combo.currentData() == "word_native"
        assert not vm.is_dirty
    finally:
        tab.close()
        set_locale(previous_locale)


def test_document_marker_control_defaults_off_preserves_choice_when_numbering_disabled(qapp) -> None:
    from docwen_gui.widgets.settings.document_tab import DocumentTab

    config = SettingsConfig()
    config.conversion_defaults.document = {
        "to_md_add_numbering": False,
        "export_numbering_markers": False,
    }
    vm = SettingsViewModel(config=config)
    tab = DocumentTab(vm)
    try:
        markers = tab._widgets["export_numbering_markers"]
        numbering = tab._widgets["to_md_add_numbering"]
        assert isinstance(markers, QCheckBox)
        assert isinstance(numbering, QCheckBox)
        assert not markers.isChecked()
        assert not markers.isEnabled()
        numbering.setChecked(True)
        assert markers.isEnabled()
        markers.setChecked(True)
        assert vm.config.conversion_defaults.document["export_numbering_markers"] is True
        numbering.setChecked(False)
        assert not markers.isEnabled()
        assert markers.isChecked()
        vm.cancel()
        tab.reload_from_config()
        assert not markers.isChecked()
        assert not markers.isEnabled()
    finally:
        tab.close()
