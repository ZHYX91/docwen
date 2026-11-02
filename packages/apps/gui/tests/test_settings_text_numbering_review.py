"""General heading-number review controls remain reachable in every UI language."""

from __future__ import annotations

import pytest
from PySide6.QtWidgets import QWidget

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
        assert card.isAncestorOf(tab._numbering_text)
        assert card.isAncestorOf(tab._numbering_comments)
        assert tab._numbering_text.isChecked()
        assert tab._numbering_comments.isChecked()
        assert tab._numbering_comments.isEnabled()
        assert tab._numbering_text.toolTip() == t("settings.text.numbering_text_help")
        assert tab._numbering_comments.toolTip() == t("settings.text.numbering_comments_help")
        assert "settings." not in tab._numbering_text.accessibleName()

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
        assert vm.config.text.heading_numbering_render_mode == "word_native"

        tab._numbering_text.setChecked(False)
        vm.cancel()
        assert tab._numbering_text.isChecked()
        assert tab._numbering_comments.isChecked()
        assert tab._render_mode_combo.currentData() == "word_native"
        assert not vm.is_dirty
    finally:
        tab.close()
        set_locale(previous_locale)
