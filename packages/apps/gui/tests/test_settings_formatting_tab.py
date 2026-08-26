from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


def _combo_values(combo) -> list[object]:
    return [combo.itemData(i) for i in range(combo.count())]


def _combo_texts(combo) -> list[str]:
    return [combo.itemText(i) for i in range(combo.count())]


def test_formatting_tab_updates_combo_fields_including_indent(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = FormattingTab(vm)

    assert _combo_values(tab._body_format) == ["preserve", "discard"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._indent_spaces) == [2, 4]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._heading_merge_mode) == ["punct_required", "always", "never"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._table_style_mode) == ["builtin", "custom"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._builtin_table_style) == ["three_line_table", "table_grid"]  # pyright: ignore[reportPrivateUsage]

    tab.set_combo_data(tab._body_format, "discard")  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(tab._indent_spaces, 2)  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(tab._heading_merge_mode, "always")  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(tab._table_style_mode, "custom")  # pyright: ignore[reportPrivateUsage]
    tab._custom_table_style_name.setText("My Table")  # pyright: ignore[reportPrivateUsage]

    assert vm.config.formatting.body_format == "discard"
    assert vm.config.formatting.indent_spaces == 2
    assert vm.config.formatting.heading_merge_mode == "always"
    assert vm.config.formatting.table_style_mode == "custom"
    assert vm.config.formatting.custom_table_style_name == "My Table"
    assert tab._builtin_table_style.isEnabled() is False  # pyright: ignore[reportPrivateUsage]
    assert tab._custom_table_style_name.isEnabled() is True  # pyright: ignore[reportPrivateUsage]


def test_formatting_tab_preserves_yaml_list_separator_exactly(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = FormattingTab(vm)

    assert tab._list_separator.text() == "、"  # pyright: ignore[reportPrivateUsage]

    tab._list_separator.setText(", ")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.formatting.list_separator == ", "

    tab._list_separator.setText("")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.formatting.list_separator == ""


def test_formatting_tab_heading_merge_punctuation_is_exact_and_mode_dependent(qapp) -> None:
    from docwen_gui.models.settings_config import (
        DEFAULT_HEADING_MERGE_PUNCTUATION,
        FormattingConfig,
        SettingsConfig,
    )
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab

    vm = SettingsViewModel(config=SettingsConfig(formatting=FormattingConfig()))
    tab = FormattingTab(vm)

    assert DEFAULT_HEADING_MERGE_PUNCTUATION == "。：！？.:!?"
    assert tab._heading_merge_punctuation.text() == DEFAULT_HEADING_MERGE_PUNCTUATION  # pyright: ignore[reportPrivateUsage]
    assert tab._heading_merge_punctuation.isEnabled() is True  # pyright: ignore[reportPrivateUsage]

    tab._heading_merge_punctuation.setText("")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.formatting.heading_merge_punctuation == ""

    tab.set_combo_data(tab._heading_merge_mode, "always")  # pyright: ignore[reportPrivateUsage]
    assert tab._heading_merge_punctuation.isEnabled() is False  # pyright: ignore[reportPrivateUsage]
    assert vm.config.formatting.heading_merge_punctuation == ""

    tab.set_combo_data(tab._heading_merge_mode, "punct_required")  # pyright: ignore[reportPrivateUsage]
    assert tab._heading_merge_punctuation.isEnabled() is True  # pyright: ignore[reportPrivateUsage]
    assert tab._heading_merge_punctuation.text() == ""  # pyright: ignore[reportPrivateUsage]


def test_formatting_tab_table_style_controls_use_existing_locale_keys(qapp) -> None:
    from docwen_gui.i18n import get_locale, set_locale, t
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab

    previous_locale = get_locale()
    set_locale("zh_CN")
    try:
        tab = FormattingTab(SettingsViewModel(config=SettingsConfig()))

        assert _combo_texts(tab._table_style_mode) == [  # pyright: ignore[reportPrivateUsage]
            t("settings.formatting.builtin_style_radio"),
            t("settings.formatting.custom_style_radio"),
        ]
        assert _combo_texts(tab._builtin_table_style) == [  # pyright: ignore[reportPrivateUsage]
            t("settings.formatting.table_styles.three_line_table"),
            t("settings.formatting.table_styles.table_grid"),
        ]
        assert tab._table_style_mode.toolTip() == t("settings.formatting.table_style_tooltip")  # pyright: ignore[reportPrivateUsage]
        assert tab._custom_table_style_name.toolTip() == t("settings.formatting.table_style_tooltip")  # pyright: ignore[reportPrivateUsage]
    finally:
        set_locale(previous_locale)


def test_formatting_tab_user_edits_update_all_view_model_fields(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = FormattingTab(vm)

    cases = [
        ("_body_format", "body_format", "discard"),
        ("_heading_format", "heading_format", "preserve"),
        ("_table_header_format", "table_header_format", "preserve"),
        ("_page_break", "page_break_sep", "ignore"),
        ("_section_break", "section_break_sep", "---"),
        ("_horizontal_rule", "horizontal_rule_sep", "***"),
        ("_bold_syntax", "bold_syntax", "underscore"),
        ("_italic_syntax", "italic_syntax", "underscore"),
        ("_strike_syntax", "strikethrough_syntax", "html"),
        ("_highlight_syntax", "highlight_syntax", "html"),
        ("_super_syntax", "superscript_syntax", "extended"),
        ("_sub_syntax", "subscript_syntax", "extended"),
        ("_ul_syntax", "unordered_list_syntax", "plus"),
        ("_indent_spaces", "indent_spaces", 2),
        ("_md_body_format", "md_body_format", "apply"),
        ("_md_heading_format", "md_heading_format", "keep"),
        ("_md_table_header_format", "md_table_header_format", "remove"),
        ("_heading_merge_mode", "heading_merge_mode", "never"),
        ("_table_style_mode", "table_style_mode", "builtin"),
        ("_builtin_table_style", "builtin_table_style", "table_grid"),
        ("_dash_sep", "dash_sep", "ignore"),
        ("_asterisk_sep", "asterisk_sep", "page_break"),
        ("_underscore_sep", "underscore_sep", "section_break"),
    ]

    for widget_name, field_name, value in cases:
        combo = getattr(tab, widget_name)
        tab.set_combo_data(combo, value)
        assert getattr(vm.config.formatting, field_name) == value

    tab._custom_table_style_name.setText("Research Table")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.formatting.custom_table_style_name == "Research Table"

    tab._heading_merge_punctuation.setText("：§")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.formatting.heading_merge_punctuation == "：§"
