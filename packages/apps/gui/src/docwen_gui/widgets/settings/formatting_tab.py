"""Markdown syntax settings — extensions, syntax and separator mappings."""

from __future__ import annotations

from typing import cast as _cast

from PySide6.QtWidgets import QCheckBox, QComboBox, QPushButton

from docwen_core.markdown_extensions import EXTENSION_NAMES

from ...i18n import t
from ...view_models.settings_vm import SECTION_FORMATTING, SettingsViewModel
from .base_tab import BaseSettingsTab


def _sep_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.separators.ignore", "Ignore"), "ignore"),
        (t("settings.formatting.separators.dash", "--- (Dash)"), "---"),
        (t("settings.formatting.separators.asterisk", "*** (Asterisk)"), "***"),
        (t("settings.formatting.separators.underscore", "___ (Underscore)"), "___"),
    ]


def _word_sep_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.separators.ignore", "Ignore"), "ignore"),
        (t("settings.formatting.separators.page_break", "Page Break"), "page_break"),
        (t("settings.formatting.separators.section_break", "Section Break"), "section_break"),
        (t("settings.formatting.separators.horizontal_rule_1", "Horizontal Rule 1"), "horizontal_rule_1"),
        (t("settings.formatting.separators.horizontal_rule_2", "Horizontal Rule 2"), "horizontal_rule_2"),
        (t("settings.formatting.separators.horizontal_rule_3", "Horizontal Rule 3"), "horizontal_rule_3"),
    ]


def _syntax_bold_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.syntax.bold_asterisk", "* (Asterisk)"), "asterisk"),
        (t("settings.formatting.syntax.bold_underscore", "_ (Underscore)"), "underscore"),
    ]


def _syntax_italic_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.syntax.italic_asterisk", "* (Asterisk)"), "asterisk"),
        (t("settings.formatting.syntax.italic_underscore", "_ (Underscore)"), "underscore"),
    ]


def _syntax_ext_html_options(kind: str) -> list[tuple[str, str]]:
    return [
        (t(f"settings.formatting.syntax.{kind}_extended", "Extended"), "extended"),
        (t(f"settings.formatting.syntax.{kind}_html", "HTML"), "html"),
    ]


def _syntax_ul_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.syntax.unordered_dash", "Dash (-)"), "dash"),
        (t("settings.formatting.syntax.unordered_asterisk", "* (Asterisk)"), "asterisk"),
        (t("settings.formatting.syntax.unordered_plus", "+ (Plus)"), "plus"),
    ]


def _syntax_indent_options() -> list[tuple[str, int]]:
    return [
        (t("settings.formatting.syntax.indent_2_spaces", "2 spaces"), 2),
        (t("settings.formatting.syntax.indent_4_spaces", "4 spaces"), 4),
    ]


class FormattingTab(BaseSettingsTab):
    """Formatting settings tab backed by typed draft state."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._extension_checks: dict[tuple[str, str], QCheckBox] = {}
        # All combo refs — initialized in _create_interface()
        self._page_break: QComboBox = _cast(QComboBox, None)
        self._section_break: QComboBox = _cast(QComboBox, None)
        self._horizontal_rule: QComboBox = _cast(QComboBox, None)
        self._bold_syntax: QComboBox = _cast(QComboBox, None)
        self._italic_syntax: QComboBox = _cast(QComboBox, None)
        self._strike_syntax: QComboBox = _cast(QComboBox, None)
        self._highlight_syntax: QComboBox = _cast(QComboBox, None)
        self._super_syntax: QComboBox = _cast(QComboBox, None)
        self._sub_syntax: QComboBox = _cast(QComboBox, None)
        self._ul_syntax: QComboBox = _cast(QComboBox, None)
        self._indent_spaces: QComboBox = _cast(QComboBox, None)
        self._dash_sep: QComboBox = _cast(QComboBox, None)
        self._asterisk_sep: QComboBox = _cast(QComboBox, None)
        self._underscore_sep: QComboBox = _cast(QComboBox, None)
        super().__init__()
        self._load_values()

    def _create_interface(self) -> None:
        self.set_tab_description(
            t(
                "settings.markdown_extensions.description",
                "Choose which Markdown extensions to recognize and generate. Your source files are never rewritten.",
            )
        )
        for direction in ("input", "output"):
            _card, form = self.add_settings_card(
                t(f"settings.markdown_extensions.{direction}", direction.title()),
                t(f"settings.markdown_extensions.{direction}_hint", "Extensions are optional."),
                object_name=f"markdownExtensions{direction.title()}Card",
            )
            for name in EXTENSION_NAMES:
                checkbox = self.create_checkbox(t(f"settings.markdown_extensions.{name}", name.replace("_", " ")))
                checkbox.setObjectName(f"markdownExtension{direction.title()}{name.title().replace('_', '')}")
                self._extension_checks[(direction, name)] = checkbox
                self.add_form_row(form, "", checkbox)
                checkbox.toggled.connect(self._save_extensions)
            preset = QPushButton(t("settings.markdown_extensions.obsidian_preset", "Use Obsidian extensions"))
            preset.setObjectName(f"markdownExtensions{direction.title()}Preset")
            preset.clicked.connect(lambda _checked=False, selected=direction: self._apply_extension_preset(selected))
            form.addRow(preset)
        # ── DOCX → MD: Separator Mapping ───────────────────────────────
        _c2, f2 = self.add_settings_card(
            f"{t('settings.formatting.docx_to_md_section', 'DOCX to MD')} — {t('settings.formatting.separator_mapping', 'Separator Mapping')}",
            t("settings.formatting.separator_mapping_tooltip", "Map document separators to Markdown equivalents."),
            object_name="formattingDocxSeparatorsCard",
        )
        self._page_break = self.create_combobox(
            _sep_options(), t("settings.formatting.page_break_tooltip", "Page break mapping")
        )
        self.add_form_row(f2, t("settings.formatting.page_break_label", "Page Break:"), self._page_break)
        self._section_break = self.create_combobox(
            _sep_options(), t("settings.formatting.section_break_tooltip", "Section break mapping")
        )
        self.add_form_row(f2, t("settings.formatting.section_break_label", "Section Break:"), self._section_break)
        self._horizontal_rule = self.create_combobox(
            _sep_options(), t("settings.formatting.horizontal_rule_tooltip", "Horizontal rule mapping")
        )
        self.add_form_row(f2, t("settings.formatting.horizontal_rule_label", "Horizontal Rule:"), self._horizontal_rule)

        # ── DOCX → MD: Syntax ──────────────────────────────────────────
        _c3, f3 = self.add_settings_card(
            f"{t('settings.formatting.docx_to_md_section', 'DOCX to MD')} — {t('settings.formatting.syntax_selection', 'Syntax Selection')}",
            t(
                "settings.formatting.syntax_selection_tooltip",
                "Choose Markdown syntax variants for formatting elements.",
            ),
            object_name="formattingDocxSyntaxCard",
        )
        self._bold_syntax = self.create_combobox(
            _syntax_bold_options(), t("settings.formatting.bold_syntax_tooltip", "Bold syntax")
        )
        self.add_form_row(f3, t("settings.formatting.bold_syntax_label", "Bold:"), self._bold_syntax)
        self._italic_syntax = self.create_combobox(
            _syntax_italic_options(), t("settings.formatting.italic_syntax_tooltip", "Italic syntax")
        )
        self.add_form_row(f3, t("settings.formatting.italic_syntax_label", "Italic:"), self._italic_syntax)
        self._strike_syntax = self.create_combobox(
            _syntax_ext_html_options("strikethrough"),
            t("settings.formatting.strikethrough_syntax_tooltip", "Strikethrough syntax"),
        )
        self.add_form_row(
            f3, t("settings.formatting.strikethrough_syntax_label", "Strikethrough:"), self._strike_syntax
        )
        self._highlight_syntax = self.create_combobox(
            _syntax_ext_html_options("highlight"), t("settings.formatting.highlight_syntax_tooltip", "Highlight syntax")
        )
        self.add_form_row(f3, t("settings.formatting.highlight_syntax_label", "Highlight:"), self._highlight_syntax)
        self._super_syntax = self.create_combobox(
            _syntax_ext_html_options("superscript"),
            t("settings.formatting.superscript_syntax_tooltip", "Superscript syntax"),
        )
        self.add_form_row(f3, t("settings.formatting.superscript_syntax_label", "Superscript:"), self._super_syntax)
        self._sub_syntax = self.create_combobox(
            _syntax_ext_html_options("subscript"), t("settings.formatting.subscript_syntax_tooltip", "Subscript syntax")
        )
        self.add_form_row(f3, t("settings.formatting.subscript_syntax_label", "Subscript:"), self._sub_syntax)
        self._ul_syntax = self.create_combobox(
            _syntax_ul_options(), t("settings.formatting.unordered_list_tooltip", "Unordered list syntax")
        )
        self.add_form_row(f3, t("settings.formatting.unordered_list_label", "Unordered List:"), self._ul_syntax)
        self._indent_spaces = self.create_combobox(
            _syntax_indent_options(), t("settings.formatting.indent_spaces_tooltip", "Indent spaces")
        )
        self.add_form_row(f3, t("settings.formatting.indent_spaces_label", "Indent:"), self._indent_spaces)

        # ── MD → DOCX: Separator Mapping ───────────────────────────────
        _c5, f5 = self.add_settings_card(
            f"{t('settings.formatting.md_to_docx_section', 'MD to DOCX')} — {t('settings.formatting.md_separator_mapping', 'Separator Mapping')}",
            t("settings.formatting.md_separator_mapping_tooltip", "Map Markdown separators to document equivalents."),
            object_name="formattingMdSeparatorsCard",
        )
        self._dash_sep = self.create_combobox(
            _word_sep_options(), t("settings.formatting.dash_tooltip", "Dash separator mapping")
        )
        self.add_form_row(f5, t("settings.formatting.dash_label", "Dash (---):"), self._dash_sep)
        self._asterisk_sep = self.create_combobox(
            _word_sep_options(), t("settings.formatting.asterisk_tooltip", "Asterisk separator mapping")
        )
        self.add_form_row(f5, t("settings.formatting.asterisk_label", "Asterisk (***):"), self._asterisk_sep)
        self._underscore_sep = self.create_combobox(
            _word_sep_options(), t("settings.formatting.underscore_tooltip", "Underscore separator mapping")
        )
        self.add_form_row(f5, t("settings.formatting.underscore_label", "Underscore (___):"), self._underscore_sep)

        # Wire all combos
        self._wire_combo(self._page_break, "page_break_sep")
        self._wire_combo(self._section_break, "section_break_sep")
        self._wire_combo(self._horizontal_rule, "horizontal_rule_sep")
        self._wire_combo(self._bold_syntax, "bold_syntax")
        self._wire_combo(self._italic_syntax, "italic_syntax")
        self._wire_combo(self._strike_syntax, "strikethrough_syntax")
        self._wire_combo(self._highlight_syntax, "highlight_syntax")
        self._wire_combo(self._super_syntax, "superscript_syntax")
        self._wire_combo(self._sub_syntax, "subscript_syntax")
        self._wire_combo(self._ul_syntax, "unordered_list_syntax")
        self._wire_combo(self._indent_spaces, "indent_spaces")
        self._wire_combo(self._dash_sep, "dash_sep")
        self._wire_combo(self._asterisk_sep, "asterisk_sep")
        self._wire_combo(self._underscore_sep, "underscore_sep")

    def _save_extensions(self, _checked: bool = False) -> None:
        values = {
            direction: {name: self._extension_checks[(direction, name)].isChecked() for name in EXTENSION_NAMES}
            for direction in ("input", "output")
        }
        self._vm.set_field(SECTION_FORMATTING, "markdown_extensions", values)

    def _apply_extension_preset(self, direction: str) -> None:
        for (group, _name), checkbox in self._extension_checks.items():
            if group == direction:
                checkbox.blockSignals(True)
                checkbox.setChecked(True)
                checkbox.blockSignals(False)
        self._save_extensions()

    def _wire_combo(self, combo: QComboBox, key: str) -> None:
        combo.currentIndexChanged.connect(
            lambda _idx, k=key, c=combo: self._vm.set_field(SECTION_FORMATTING, k, self.get_combo_data(c))
        )

    def _load_values(self) -> None:
        fmt = self._vm.config.formatting
        for (direction, name), checkbox in self._extension_checks.items():
            checkbox.blockSignals(True)
            checkbox.setChecked(fmt.markdown_extensions[direction][name])
            checkbox.blockSignals(False)
        self.set_combo_data(self._page_break, fmt.page_break_sep)
        self.set_combo_data(self._section_break, fmt.section_break_sep)
        self.set_combo_data(self._horizontal_rule, fmt.horizontal_rule_sep)
        self.set_combo_data(self._bold_syntax, fmt.bold_syntax)
        self.set_combo_data(self._italic_syntax, fmt.italic_syntax)
        self.set_combo_data(self._strike_syntax, fmt.strikethrough_syntax)
        self.set_combo_data(self._highlight_syntax, fmt.highlight_syntax)
        self.set_combo_data(self._super_syntax, fmt.superscript_syntax)
        self.set_combo_data(self._sub_syntax, fmt.subscript_syntax)
        self.set_combo_data(self._ul_syntax, fmt.unordered_list_syntax)
        self.set_combo_data(self._indent_spaces, fmt.indent_spaces)
        self.set_combo_data(self._dash_sep, fmt.dash_sep)
        self.set_combo_data(self._asterisk_sep, fmt.asterisk_sep)
        self.set_combo_data(self._underscore_sep, fmt.underscore_sep)

    def reload_from_config(self) -> None:
        self._load_values()
