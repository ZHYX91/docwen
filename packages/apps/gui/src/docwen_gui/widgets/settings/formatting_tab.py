"""Formatting settings tab — DOCX<->MD format/separator/syntax choices."""

from __future__ import annotations

from typing import cast as _cast

from PySide6.QtWidgets import QComboBox, QLineEdit

from ...i18n import t
from ...view_models.settings_vm import SECTION_FORMATTING, SettingsViewModel
from .base_tab import BaseSettingsTab


def _fmt_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.options.preserve_format", "Preserve Formatting"), "preserve"),
        (t("settings.formatting.options.discard_format", "Discard Formatting"), "discard"),
    ]


def _md_fmt_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.options.apply_format", "Apply Formatting"), "apply"),
        (t("settings.formatting.options.keep_markup", "Keep Markup"), "keep"),
        (t("settings.formatting.options.clean_markup", "Remove Markup"), "remove"),
    ]


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


def _heading_merge_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.options.heading_merge_punct_required", "Punctuation Required"), "punct_required"),
        (t("settings.formatting.options.heading_merge_always", "Always Merge"), "always"),
        (t("settings.formatting.options.heading_merge_never", "Never Merge"), "never"),
    ]


def _table_style_mode_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.builtin_style_radio", "Use built-in style"), "builtin"),
        (t("settings.formatting.custom_style_radio", "Use custom style name"), "custom"),
    ]


def _builtin_table_style_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.table_styles.three_line_table", "Three-line Table"), "three_line_table"),
        (t("settings.formatting.table_styles.table_grid", "Table Grid"), "table_grid"),
    ]


class FormattingTab(BaseSettingsTab):
    """Formatting settings tab backed by typed draft state."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        # All combo refs — initialized in _create_interface()
        self._body_format: QComboBox = _cast(QComboBox, None)
        self._heading_format: QComboBox = _cast(QComboBox, None)
        self._table_header_format: QComboBox = _cast(QComboBox, None)
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
        self._md_body_format: QComboBox = _cast(QComboBox, None)
        self._md_heading_format: QComboBox = _cast(QComboBox, None)
        self._md_table_header_format: QComboBox = _cast(QComboBox, None)
        self._heading_merge_mode: QComboBox = _cast(QComboBox, None)
        self._heading_merge_punctuation: QLineEdit = _cast(QLineEdit, None)
        self._list_separator: QLineEdit = _cast(QLineEdit, None)
        self._table_style_mode: QComboBox = _cast(QComboBox, None)
        self._builtin_table_style: QComboBox = _cast(QComboBox, None)
        self._custom_table_style_name: QLineEdit = _cast(QLineEdit, None)
        self._dash_sep: QComboBox = _cast(QComboBox, None)
        self._asterisk_sep: QComboBox = _cast(QComboBox, None)
        self._underscore_sep: QComboBox = _cast(QComboBox, None)
        super().__init__()
        self._load_values()

    def _create_interface(self) -> None:
        # ── DOCX → MD: Format Processing ────────────────────────────────
        _c1, f1 = self.add_settings_card(
            f"{t('settings.formatting.docx_to_md_section', 'DOCX to MD')} — {t('settings.formatting.format_processing', 'Format Processing')}",
            t(
                "settings.formatting.format_processing_tooltip",
                "How to handle formatting when converting from DOCX to Markdown.",
            ),
            object_name="formattingDocxProcessingCard",
        )
        self._body_format = self.create_combobox(
            _fmt_options(), t("settings.formatting.body_format_tooltip", "How to handle body text formatting")
        )
        self.add_form_row(f1, t("settings.formatting.body_format_label", "Body Text Format:"), self._body_format)

        self._heading_format = self.create_combobox(
            _fmt_options(), t("settings.formatting.heading_format_tooltip", "How to handle heading formatting")
        )
        self.add_form_row(f1, t("settings.formatting.heading_format_label", "Heading Format:"), self._heading_format)

        self._table_header_format = self.create_combobox(
            _fmt_options(),
            t("settings.formatting.table_header_format_tooltip", "How to handle table header formatting"),
        )
        self.add_form_row(
            f1, t("settings.formatting.table_header_format_label", "Table Header Format:"), self._table_header_format
        )

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

        # ── MD → DOCX: Format Processing ───────────────────────────────
        _c4, f4 = self.add_settings_card(
            f"{t('settings.formatting.md_to_docx_section', 'MD to DOCX')} — {t('settings.formatting.md_format_processing', 'Format Processing')}",
            t(
                "settings.formatting.md_format_processing_tooltip",
                "How to handle formatting when converting from MD to DOCX.",
            ),
            object_name="formattingMdProcessingCard",
        )
        self._md_body_format = self.create_combobox(
            _md_fmt_options(), t("settings.formatting.md_body_format_tooltip", "Body format markup handling")
        )
        self.add_form_row(f4, t("settings.formatting.md_body_format_label", "Body Text Format:"), self._md_body_format)
        self._md_heading_format = self.create_combobox(
            _md_fmt_options(), t("settings.formatting.md_heading_format_tooltip", "Heading format markup handling")
        )
        self.add_form_row(
            f4, t("settings.formatting.md_heading_format_label", "Heading Format:"), self._md_heading_format
        )
        self._md_table_header_format = self.create_combobox(
            _md_fmt_options(),
            t("settings.formatting.md_table_header_format_tooltip", "Table header format markup handling"),
        )
        self.add_form_row(
            f4,
            t("settings.formatting.md_table_header_format_label", "Table Header Format:"),
            self._md_table_header_format,
        )
        self._heading_merge_mode = self.create_combobox(
            _heading_merge_options(), t("settings.formatting.heading_merge_mode_tooltip", "Heading merge mode")
        )
        self.add_form_row(
            f4, t("settings.formatting.heading_merge_mode_label", "Heading Merge Mode:"), self._heading_merge_mode
        )
        self._heading_merge_punctuation = QLineEdit(self)
        self._heading_merge_punctuation.setToolTip(
            t(
                "settings.formatting.heading_merge_punctuation_tooltip",
                "Characters that trigger heading/body merging when punctuation is required. Empty disables punctuation-triggered merging.",
            )
        )
        self.add_form_row(
            f4,
            t("settings.formatting.heading_merge_punctuation_label", "Merge punctuation:"),
            self._heading_merge_punctuation,
        )
        self._list_separator = QLineEdit(self)
        self._list_separator.setToolTip(
            t(
                "settings.formatting.list_separator_tooltip",
                "Separator used when joining YAML list values into a string.",
            )
        )
        self.add_form_row(
            f4,
            t("settings.formatting.list_separator_label", "YAML List Separator") + ":",
            self._list_separator,
        )

        # ── MD → DOCX: Table Style ────────────────────────────────────
        _c_table, f_table = self.add_settings_card(
            f"{t('settings.formatting.md_to_docx_section', 'MD to DOCX')} — {t('settings.formatting.table_style', 'Table Style')}",
            t(
                "settings.formatting.table_style_tooltip",
                "Choose the Word table style used for Markdown tables.",
            ),
            object_name="formattingMdTableStyleCard",
        )
        self._table_style_mode = self.create_combobox(
            _table_style_mode_options(),
            t("settings.formatting.table_style_tooltip", "Use a built-in table style or a custom style name."),
        )
        self.add_form_row(
            f_table,
            t("settings.formatting.table_style", "Table Style") + ":",
            self._table_style_mode,
        )
        self._builtin_table_style = self.create_combobox(
            _builtin_table_style_options(),
            t("settings.formatting.table_style_tooltip", "Built-in table style for Markdown tables."),
        )
        self.add_form_row(
            f_table,
            t("settings.formatting.builtin_style_radio", "Use built-in style") + ":",
            self._builtin_table_style,
        )
        self._custom_table_style_name = QLineEdit(self)
        self._custom_table_style_name.setToolTip(
            t("settings.formatting.table_style_tooltip", "Custom Word table style name.")
        )
        self.add_form_row(
            f_table,
            t("settings.formatting.custom_style_radio", "Use custom style name") + ":",
            self._custom_table_style_name,
        )

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
        self._wire_combo(self._body_format, "body_format")
        self._wire_combo(self._heading_format, "heading_format")
        self._wire_combo(self._table_header_format, "table_header_format")
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
        self._wire_combo(self._md_body_format, "md_body_format")
        self._wire_combo(self._md_heading_format, "md_heading_format")
        self._wire_combo(self._md_table_header_format, "md_table_header_format")
        self._wire_combo(self._heading_merge_mode, "heading_merge_mode")
        self._heading_merge_mode.currentIndexChanged.connect(
            lambda _idx: self._sync_heading_merge_punctuation_control()
        )
        self._heading_merge_punctuation.textChanged.connect(
            lambda text: self._vm.set_field(SECTION_FORMATTING, "heading_merge_punctuation", text)
        )
        self._list_separator.textChanged.connect(
            lambda text: self._vm.set_field(SECTION_FORMATTING, "list_separator", text)
        )
        self._wire_combo(self._table_style_mode, "table_style_mode")
        self._wire_combo(self._builtin_table_style, "builtin_table_style")
        self._table_style_mode.currentIndexChanged.connect(lambda _idx: self._sync_table_style_controls())
        self._custom_table_style_name.textChanged.connect(
            lambda text: self._vm.set_field(SECTION_FORMATTING, "custom_table_style_name", text.strip())
        )
        self._wire_combo(self._dash_sep, "dash_sep")
        self._wire_combo(self._asterisk_sep, "asterisk_sep")
        self._wire_combo(self._underscore_sep, "underscore_sep")

    def _wire_combo(self, combo: QComboBox, key: str) -> None:
        combo.currentIndexChanged.connect(
            lambda _idx, k=key, c=combo: self._vm.set_field(SECTION_FORMATTING, k, self.get_combo_data(c))
        )

    def _load_values(self) -> None:
        fmt = self._vm.config.formatting
        self.set_combo_data(self._body_format, fmt.body_format)
        self.set_combo_data(self._heading_format, fmt.heading_format)
        self.set_combo_data(self._table_header_format, fmt.table_header_format)
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
        self.set_combo_data(self._md_body_format, fmt.md_body_format)
        self.set_combo_data(self._md_heading_format, fmt.md_heading_format)
        self.set_combo_data(self._md_table_header_format, fmt.md_table_header_format)
        self.set_combo_data(self._heading_merge_mode, fmt.heading_merge_mode)
        self._heading_merge_punctuation.setText(fmt.heading_merge_punctuation)
        self._sync_heading_merge_punctuation_control()
        self._list_separator.setText(fmt.list_separator)
        self.set_combo_data(self._table_style_mode, fmt.table_style_mode)
        self.set_combo_data(self._builtin_table_style, fmt.builtin_table_style)
        self._custom_table_style_name.setText(fmt.custom_table_style_name)
        self._sync_table_style_controls()
        self.set_combo_data(self._dash_sep, fmt.dash_sep)
        self.set_combo_data(self._asterisk_sep, fmt.asterisk_sep)
        self.set_combo_data(self._underscore_sep, fmt.underscore_sep)

    def reload_from_config(self) -> None:
        self._load_values()

    def _sync_table_style_controls(self) -> None:
        use_builtin = self.get_combo_data(self._table_style_mode) == "builtin"
        self._builtin_table_style.setEnabled(use_builtin)
        self._custom_table_style_name.setEnabled(not use_builtin)

    def _sync_heading_merge_punctuation_control(self) -> None:
        self._heading_merge_punctuation.setEnabled(self.get_combo_data(self._heading_merge_mode) == "punct_required")
