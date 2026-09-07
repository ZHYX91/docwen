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
    from docwen_gui.widgets.settings.document_tab import DocumentTab
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab
    from docwen_gui.widgets.settings.text_tab import TextTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = FormattingTab(vm)
    text_tab = TextTab(vm)
    document_tab = DocumentTab(vm)
    text = text_tab._content_controls
    document = document_tab._content_controls

    assert _combo_values(document._body_format) == ["preserve", "discard"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._indent_spaces) == [2, 4]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(text._heading_merge_mode) == ["punct_required", "always", "never"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(text._table_style_mode) == ["builtin", "custom"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(text._builtin_table_style) == ["three_line_table", "table_grid"]  # pyright: ignore[reportPrivateUsage]

    tab.set_combo_data(document._body_format, "discard")  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(tab._indent_spaces, 2)  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(text._heading_merge_mode, "always")  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(text._table_style_mode, "custom")  # pyright: ignore[reportPrivateUsage]
    text._custom_table_style_name.setText("My Table")  # pyright: ignore[reportPrivateUsage]

    assert vm.config.document.body_format == "discard"
    assert vm.config.formatting.indent_spaces == 2
    assert vm.config.text.heading_merge_mode == "always"
    assert vm.config.text.table_style_mode == "custom"
    assert vm.config.text.custom_table_style_name == "My Table"
    assert text._builtin_table_style.isEnabled() is False  # pyright: ignore[reportPrivateUsage]
    assert text._custom_table_style_name.isEnabled() is True  # pyright: ignore[reportPrivateUsage]


def test_formatting_tab_preserves_yaml_list_separator_exactly(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.text_tab import TextTab

    vm = SettingsViewModel(config=SettingsConfig())
    text_tab = TextTab(vm)
    text = text_tab._content_controls

    assert text._list_separator.text() == "、"  # pyright: ignore[reportPrivateUsage]

    text._list_separator.setText(", ")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.text.list_separator == ", "

    text._list_separator.setText("")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.text.list_separator == ""


def test_formatting_tab_heading_merge_punctuation_is_exact_and_mode_dependent(qapp) -> None:
    from docwen_gui.models.settings_config import (
        DEFAULT_HEADING_MERGE_PUNCTUATION,
        FormattingConfig,
        SettingsConfig,
    )
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab
    from docwen_gui.widgets.settings.text_tab import TextTab

    vm = SettingsViewModel(config=SettingsConfig(formatting=FormattingConfig()))
    tab = FormattingTab(vm)
    text_tab = TextTab(vm)
    text = text_tab._content_controls

    assert DEFAULT_HEADING_MERGE_PUNCTUATION == "。：！？.:!?"
    assert text._heading_merge_punctuation.text() == DEFAULT_HEADING_MERGE_PUNCTUATION  # pyright: ignore[reportPrivateUsage]
    assert text._heading_merge_punctuation.isEnabled() is True  # pyright: ignore[reportPrivateUsage]

    text._heading_merge_punctuation.setText("")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.text.heading_merge_punctuation == ""

    tab.set_combo_data(text._heading_merge_mode, "always")  # pyright: ignore[reportPrivateUsage]
    assert text._heading_merge_punctuation.isEnabled() is False  # pyright: ignore[reportPrivateUsage]
    assert vm.config.text.heading_merge_punctuation == ""

    tab.set_combo_data(text._heading_merge_mode, "punct_required")  # pyright: ignore[reportPrivateUsage]
    assert text._heading_merge_punctuation.isEnabled() is True  # pyright: ignore[reportPrivateUsage]
    assert text._heading_merge_punctuation.text() == ""  # pyright: ignore[reportPrivateUsage]


def test_formatting_tab_table_style_controls_use_existing_locale_keys(qapp) -> None:
    from docwen_gui.i18n import get_locale, set_locale, t
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.text_tab import TextTab

    previous_locale = get_locale()
    set_locale("zh_CN")
    try:
        vm = SettingsViewModel(config=SettingsConfig())
        text_tab = TextTab(vm)
        text = text_tab._content_controls

        assert _combo_texts(text._table_style_mode) == [  # pyright: ignore[reportPrivateUsage]
            t("settings.formatting.builtin_style_radio"),
            t("settings.formatting.custom_style_radio"),
        ]
        assert _combo_texts(text._builtin_table_style) == [  # pyright: ignore[reportPrivateUsage]
            t("settings.formatting.table_styles.three_line_table"),
            t("settings.formatting.table_styles.table_grid"),
        ]
        assert text._table_style_mode.toolTip() == t("settings.formatting.table_style_tooltip")  # pyright: ignore[reportPrivateUsage]
        assert text._custom_table_style_name.toolTip() == t("settings.formatting.table_style_tooltip")  # pyright: ignore[reportPrivateUsage]
    finally:
        set_locale(previous_locale)


def test_formatting_tab_user_edits_update_all_view_model_fields(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.document_tab import DocumentTab
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab
    from docwen_gui.widgets.settings.text_tab import TextTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = FormattingTab(vm)
    text_tab = TextTab(vm)
    document_tab = DocumentTab(vm)
    text = text_tab._content_controls
    document = document_tab._content_controls

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
        text_fields = {
            "md_body_format",
            "md_heading_format",
            "md_table_header_format",
            "heading_merge_mode",
            "table_style_mode",
            "builtin_table_style",
        }
        document_fields = {"body_format", "heading_format", "table_header_format"}
        owner, config = (
            (text, vm.config.text)
            if field_name in text_fields
            else (document, vm.config.document)
            if field_name in document_fields
            else (tab, vm.config.formatting)
        )
        combo = getattr(owner, widget_name)
        tab.set_combo_data(combo, value)
        config = (
            vm.config.text
            if field_name in text_fields
            else vm.config.document
            if field_name in document_fields
            else vm.config.formatting
        )
        assert getattr(config, field_name) == value

    text._custom_table_style_name.setText("Research Table")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.text.custom_table_style_name == "Research Table"

    text._heading_merge_punctuation.setText("：§")  # pyright: ignore[reportPrivateUsage]
    assert vm.config.text.heading_merge_punctuation == "：§"


def test_extension_controls_and_directional_preset_update_only_the_selected_direction(qapp) -> None:
    from PySide6.QtWidgets import QCheckBox, QPushButton

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = FormattingTab(vm)
    control = tab.findChild(QCheckBox, "markdownExtensionInputStructuralTables")
    preset = tab.findChild(QPushButton, "markdownExtensionsOutputPreset")
    assert control is not None and preset is not None
    assert not control.isChecked()
    control.setChecked(True)
    assert vm.config.formatting.markdown_extensions["input"] == {
        "structural_tables": True,
        "captions_references": False,
        "extended_headings": False,
        "typed_endnotes": False,
    }
    assert not any(vm.config.formatting.markdown_extensions["output"].values())
    preset.click()
    assert all(vm.config.formatting.markdown_extensions["output"].values())
    assert not vm.config.formatting.markdown_extensions["input"]["captions_references"]
    vm.cancel_changes()
    tab.reload_from_config()
    assert not control.isChecked()
    assert not any(vm.config.formatting.markdown_extensions["output"].values())


@pytest.mark.parametrize("locale", ["zh_CN", "en_US"])
@pytest.mark.parametrize("font_size", [12, 15])
def test_extension_labels_use_available_row_width_and_remain_accessible(qapp, locale, font_size) -> None:
    from PySide6.QtGui import QFont
    from PySide6.QtWidgets import QCheckBox, QWidget

    from docwen_gui.i18n import get_locale, set_locale
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.check_box import SettingsCheckBox
    from docwen_gui.widgets.settings.formatting_tab import FormattingTab

    previous = get_locale()
    set_locale(locale)
    tab = FormattingTab(SettingsViewModel(config=SettingsConfig()))
    tab.setFont(QFont("Microsoft YaHei", font_size))
    tab.resize(600, 850)
    try:
        tab.show()
        for _ in range(8):
            qapp.processEvents()
        for direction in ("Input", "Output"):
            card = tab.findChild(QWidget, f"markdownExtensions{direction}Card")
            assert card is not None
            controls = card.findChildren(SettingsCheckBox)
            assert len(controls) == 4
            assert len({control.geometry().left() for control in controls}) == 1
            for control in controls:
                assert control.text() and control.accessibleName() == control.text()
                assert control.height() >= control.heightForWidth(control.width())
                painted_text = QCheckBox.text(control)
                assert "".join(painted_text.split()) == "".join(control.text().split())
                assert not control.font().bold()
    finally:
        tab.close()
        set_locale(previous)
