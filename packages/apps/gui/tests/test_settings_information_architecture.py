"""Settings grouping must not narrow Markdown policy or change config ownership."""

from __future__ import annotations

from copy import deepcopy

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QCheckBox, QComboBox, QLabel, QWidget

from docwen_gui.i18n import get_locale, set_locale, t
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings import dialog as settings_dialog
from docwen_gui.widgets.settings.dialog import NAVIGATION_GROUPS, TAB_KEYS, SettingsDialog
from docwen_gui.widgets.settings.export_tab import ExportTab
from docwen_gui.widgets.settings.formatting_tab import FormattingTab, MarkdownInputTab

pytestmark = pytest.mark.gui


def test_groups_have_unique_pages_without_creating_more_settings() -> None:
    assert [(group.key, len(group.page_keys)) for group in NAVIGATION_GROUPS] == [
        ("general", 5),
        ("inputs", 6),
        ("content", 5),
        ("optimizations", 2),
    ]
    assert len(TAB_KEYS) == len(set(TAB_KEYS)) == 18
    assert set(TAB_KEYS) == {
        "general",
        "output",
        "templates",
        "software",
        "logging",
        "text",
        "document",
        "spreadsheet",
        "image",
        "layout",
        "other",
        "formatting",
        "markdown_input",
        "link",
        "export",
        "proofread",
        "gongwen",
        "invoice",
    }


def test_markdown_output_and_input_precede_word_specific_mappings(qapp) -> None:
    vm = SettingsViewModel()
    before = deepcopy(vm.config)
    page = FormattingTab(vm)
    input_page = MarkdownInputTab(vm)
    try:
        container = page.findChild(QWidget, "settingsTabScrollContainer")
        assert container is not None
        layout = container.layout()
        assert layout is not None
        names = []
        for index in range(layout.count()):
            item = layout.itemAt(index)
            widget = item.widget() if item is not None else None
            if widget is not None:
                names.append(widget.objectName())
        expected = [
            "markdownExtensionsOutputCard",
            "formattingOutputSyntaxCard",
            "formattingDocxSeparatorsCard",
        ]
        assert [name for name in names if name in expected] == expected
        assert page.findChild(QWidget, "markdownExtensionsInputCard") is None
        assert input_page.findChild(QWidget, "markdownExtensionsOutputCard") is None
        assert input_page.findChild(QWidget, "formattingMdMermaidCard") is not None
        syntax = page.findChild(QWidget, "formattingOutputSyntaxCard")
        assert syntax is not None
        help_label = syntax.findChild(QLabel, "settingsCardDescription")
        assert help_label is not None
        assert help_label.text() == t("settings.formatting.syntax_selection_tooltip")
        assert vm.config == before
    finally:
        page.close()
        input_page.close()


def test_extension_directions_remain_independent_after_card_reordering(qapp) -> None:
    vm = SettingsViewModel()
    page = FormattingTab(vm)
    input_page = MarkdownInputTab(vm)
    try:
        original_input = deepcopy(vm.config.formatting.markdown_extensions["input"])
        original_other = deepcopy(vm.config.conversion_defaults)
        output_control = page.findChild(QCheckBox, "markdownExtensionOutputStructuralTables")
        input_control = input_page.findChild(QCheckBox, "markdownExtensionInputStructuralTables")
        assert output_control is not None and input_control is not None
        output_control.setChecked(not output_control.isChecked())
        assert vm.config.formatting.markdown_extensions["output"]["structural_tables"] == output_control.isChecked()
        assert vm.config.formatting.markdown_extensions["input"] == original_input
        selected_output = deepcopy(vm.config.formatting.markdown_extensions["output"])
        input_control.setChecked(not input_control.isChecked())
        assert vm.config.formatting.markdown_extensions["output"] == selected_output
        assert vm.config.conversion_defaults == original_other
    finally:
        page.close()
        input_page.close()


def test_ocr_page_explains_shared_scope_without_duplicating_input_defaults(qapp) -> None:
    vm = SettingsViewModel()
    before = deepcopy(vm.config.conversion_defaults)
    page = ExportTab(vm)
    try:
        descriptions = [label.text() for label in page.findChildren(QLabel)]
        assert t("settings.export.description") in descriptions
        controls = page.findChildren(QComboBox)
        assert controls
        controls[0].setCurrentIndex((controls[0].currentIndex() + 1) % controls[0].count())
        assert vm.config.conversion_defaults == before
    finally:
        page.close()


@pytest.mark.parametrize(
    "locale",
    ["zh_CN", "zh_TW", "en_US", "de_DE", "fr_FR", "es_ES", "pt_BR", "ru_RU", "ja_JP", "ko_KR", "vi_VN", "ar_SA"],
)
def test_groups_and_short_ocr_label_are_localized(qapp, monkeypatch, locale) -> None:
    previous = get_locale()
    set_locale(locale)
    title_keys = {
        "templates": "settings.templates.title",
        "gongwen": "settings.gongwen.title",
        "invoice": "settings.invoice.title",
    }
    monkeypatch.setattr(
        settings_dialog, "TAB_NAMES", {key: t(title_keys.get(key, f"settings.tabs.{key}")) for key in TAB_KEYS}
    )
    window = None
    try:
        window = SettingsDialog(view_model=SettingsViewModel())
        for group in NAVIGATION_GROUPS:
            label = window.findChild(QLabel, f"settingsNavGroup_{group.key}")
            assert label is not None
            assert label.text() == t(f"settings.navigation.{group.key}")
            assert not label.text().startswith("settings.")
            assert label.focusPolicy() == Qt.FocusPolicy.NoFocus
        if locale == "zh_CN":
            assert t("settings.tabs.export") == "OCR 与图片"
        assert "Markdown" not in t("settings.tabs.export")
        assert window.activate_section("export")
        assert window.current_section() == "export"
    finally:
        if window is not None:
            window.close()
        set_locale(previous)


def test_chinese_default_width_keeps_full_sidebar(qapp, monkeypatch) -> None:
    from PySide6.QtWidgets import QScrollArea

    previous = get_locale()
    set_locale("zh_CN")
    title_keys = {
        "templates": "settings.templates.title",
        "gongwen": "settings.gongwen.title",
        "invoice": "settings.invoice.title",
    }
    monkeypatch.setattr(
        settings_dialog, "TAB_NAMES", {key: t(title_keys.get(key, f"settings.tabs.{key}")) for key in TAB_KEYS}
    )
    window = None
    try:
        window = SettingsDialog(view_model=SettingsViewModel())
        window.resize(700, 600)
        window.show()
        qapp.processEvents()
        assert window.activate_section("export")
        qapp.processEvents()
        navigation = window.findChild(QScrollArea, "settingsNavigation")
        selector = window.findChild(QComboBox, "settingsPageSelector")
        assert navigation is not None and selector is not None
        assert window.width() == 700
        assert navigation.isVisible()
        assert selector.isHidden()
        assert t("settings.tabs.export") == "OCR 与图片"
    finally:
        if window is not None:
            window.close()
        set_locale(previous)


@pytest.mark.parametrize(
    "previous_page,next_page",
    [
        ("logging", "text"),
        ("other", "formatting"),
        ("proofread", "gongwen"),
    ],
)
def test_keyboard_crosses_each_navigation_group(qapp, previous_page, next_page) -> None:
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QPushButton

    window = SettingsDialog(view_model=SettingsViewModel())
    try:
        window.resize(1100, 650)
        window.show()
        window.activateWindow()
        qapp.processEvents()
        assert window.activate_section(previous_page)
        first = window.findChild(QPushButton, f"settingsNavItem_{previous_page}")
        second = window.findChild(QPushButton, f"settingsNavItem_{next_page}")
        assert first is not None and second is not None
        first.setFocus(Qt.FocusReason.TabFocusReason)
        QTest.keyClick(first, Qt.Key.Key_Down)
        assert window.current_section() == next_page
        assert qapp.focusWidget() is second
        QTest.keyClick(second, Qt.Key.Key_Up)
        assert window.current_section() == previous_page
        assert qapp.focusWidget() is first
    finally:
        window.close()
