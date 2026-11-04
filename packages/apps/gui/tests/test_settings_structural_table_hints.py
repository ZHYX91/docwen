"""Structural output policy is visible without mutating ordinary merge preferences."""

from __future__ import annotations

import tomllib
from copy import deepcopy
from pathlib import Path

import pytest
from PySide6.QtCore import QEvent
from PySide6.QtWidgets import QApplication, QFrame, QLabel, QMenu, QPushButton, QToolButton

from docwen_gui.i18n import get_locale, set_locale, t
from docwen_gui.models.settings_config import SettingsConfig
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.document_tab import DocumentTab
from docwen_gui.widgets.settings.formatting_tab import FormattingTab, MarkdownInputTab
from docwen_gui.widgets.settings.spreadsheet_tab import SpreadsheetTab

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("choice", ["fill", "empty"])
def test_structural_status_tracks_output_only_without_mutating_merge_choice(qapp, choice) -> None:
    vm = SettingsViewModel(config=SettingsConfig())
    for category in ("document", "spreadsheet"):
        vm.set_conversion_default(category, "to_md_table_merge_export_strategy", choice)
    before = deepcopy(vm.config.conversion_defaults)
    tabs = [DocumentTab(vm), SpreadsheetTab(vm), FormattingTab(vm)]
    input_tab = MarkdownInputTab(vm)
    output = tabs[-1]
    labels = [tab.findChild(QLabel, "tableExportStatus") for tab in tabs]
    assert all(label is not None for label in labels)
    assert input_tab.findChild(QLabel, "tableExportStatus") is None

    for enabled in (False, True, False, True):
        output._extension_checks[("output", "structural_tables")].setChecked(enabled)
        expected = t("settings.table_export.structural_active" if enabled else "settings.table_export.ordinary_active")
        assert all(label.text() == expected for label in labels)
        input_tab._extension_checks[("input", "structural_tables")].setChecked(not enabled)
        assert all(label.text() == expected for label in labels)
        for tab in tabs[:2]:
            combo = tab.widgets["to_md_table_merge_export_strategy"]
            assert combo.currentData() == choice
            assert combo.isEnabled()
            assert "table_export_status" not in tab.collect_values()
        assert vm.config.conversion_defaults == before

    # A reload (including Reset/Cancel) must refresh the read-only projection.
    vm._config.formatting.markdown_extensions["output"]["structural_tables"] = False
    vm.config_reloaded.emit()
    assert all(label.text() == t("settings.table_export.ordinary_active") for label in labels)
    output.reload_from_config()
    output.findChild(QPushButton, "markdownExtensionsOutputPreset").click()
    assert all(label.text() == t("settings.table_export.structural_active") for label in labels)
    assert vm.config.conversion_defaults == before


def test_opening_tabs_does_not_mark_settings_dirty(qapp) -> None:
    vm = SettingsViewModel(config=SettingsConfig())
    before = deepcopy(vm.config)
    tabs = [DocumentTab(vm), SpreadsheetTab(vm), FormattingTab(vm)]
    assert len(tabs) == 3
    assert vm.config == before
    assert not vm.is_dirty


@pytest.mark.parametrize(
    "locale", sorted(path.stem for path in (Path(__file__).resolve().parents[4] / "i18n" / "locales").glob("*.toml"))
)
def test_info_buttons_preserve_multiline_help_and_exception(qapp, locale) -> None:
    previous = get_locale()
    set_locale(locale)
    try:
        vm = SettingsViewModel(config=SettingsConfig())
        for tab in (DocumentTab(vm), SpreadsheetTab(vm), FormattingTab(vm)):
            key = (
                "settings.markdown_extensions.output_help.structural_tables"
                if isinstance(tab, FormattingTab)
                else "settings.table_export.merge_strategy_help"
            )
            text = t(key)
            buttons = [button for button in tab.findChildren(QToolButton) if button.toolTip() == text]
            assert buttons
            assert text.count("\n") >= 3
            assert "<" in text and "^" in text
            QApplication.sendEvent(buttons[0], QEvent(QEvent.Type.ToolTip))
            hover = buttons[0].findChild(QFrame, "settingsHelpPopup")
            assert hover is not None
            hover_text = hover.findChild(QLabel, "settingsHelpPopupText")
            assert hover_text is not None and hover_text.text() == text
            QApplication.sendEvent(buttons[0], QEvent(QEvent.Type.Leave))
            buttons[0].click()
            popup = buttons[0].findChild(QLabel, "settingsHelpPopupText")
            assert popup is not None and popup.text() == text
            menu = buttons[0].findChild(QMenu, "settingsHelpPopup")
            assert menu is not None
            menu.close()
    finally:
        set_locale(previous)


def test_all_locale_catalogs_have_multiline_policy_and_status() -> None:
    root = Path(__file__).resolve().parents[4]
    for path in (root / "i18n" / "locales").glob("*.toml"):
        catalog = tomllib.loads(path.read_text(encoding="utf-8"))["settings"]
        help_text = catalog["table_export"]["merge_strategy_help"]
        assert help_text.count("\n") >= 5, path.name
        assert "<" in help_text and "^" in help_text
        labels = catalog["table_export"]["strategies"]
        output_help = catalog["markdown_extensions"]["output_help"]["structural_tables"]
        assert "<" in output_help and "^" in output_help
        if path.stem == "ar_SA":
            for text in (help_text, output_help):
                assert "\u2066<\u2069" in text and "\u2066^\u2069" in text
        for label in labels.values():
            assert label in help_text and label in output_help
            assert label in catalog["table_export"]["structural_active"]
            assert label in catalog["table_export"]["ordinary_active"]
        assert catalog["tabs"]["formatting"] in help_text
        assert catalog["gongwen"]["title"] in help_text
        assert catalog["gongwen"]["title"] in output_help
        assert "Gongwen" not in help_text and "Gongwen" not in output_help
        assert catalog["table_export"]["structural_active"]
        assert catalog["table_export"]["ordinary_active"]
        assert catalog["markdown_extensions"]["output_help"]["structural_tables"].count("\n") >= 4
