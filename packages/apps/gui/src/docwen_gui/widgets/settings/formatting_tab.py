"""Markdown syntax settings — extensions, syntax and separator mappings."""

from __future__ import annotations

from typing import Literal
from typing import cast as _cast

from PySide6.QtWidgets import QCheckBox, QComboBox, QPushButton

from docwen_core.markdown_extensions import EXTENSION_NAMES

from ...i18n import t
from ...view_models.settings_vm import SECTION_FORMATTING, SettingsViewModel
from .base_tab import BaseSettingsTab
from .mermaid_controls import MermaidControls


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


def _mermaid_options() -> list[tuple[str, str]]:
    return [
        (t("settings.formatting.mermaid_code", "Keep as code block"), "code"),
        (t("settings.formatting.mermaid_image", "Render as image"), "image"),
    ]


_EXTENSION_HELP_FALLBACKS: dict[tuple[str, str], str] = {
    ("input", "structural_tables"): (
        "Recognize Structural Tables markers such as < and ^ and preserve extended table structure. "
        "When disabled, treat them as ordinary Markdown table content."
    ),
    ("output", "structural_tables"): (
        "Write Structural Tables markers and extended structure when needed. When disabled, flatten to ordinary "
        "Markdown tables and report lost structure."
    ),
    ("input", "captions_references"): (
        "Recognize Number Suite captions, identifiers and cross-references. When disabled, declarations and @ "
        "references remain visible text."
    ),
    ("output", "captions_references"): (
        "Generate Number Suite captions and cross-references when available. When disabled, emit ordinary text and "
        "report lost references."
    ),
    ("input", "extended_headings"): (
        "Recognize H7–H9 syntax. When disabled, seven to nine leading # characters remain visible text."
    ),
    ("output", "extended_headings"): (
        "Generate H7–H9 headings. When disabled, those headings become H6 and the loss is reported."
    ),
    ("input", "typed_endnotes"): (
        "Use note labels such as endnote: to distinguish endnotes from footnotes. When disabled, those labels are "
        "treated as ordinary footnotes."
    ),
    ("output", "typed_endnotes"): (
        "Preserve the footnote/endnote distinction with typed note labels. When disabled, endnotes are exported as "
        "ordinary footnotes and the loss is reported."
    ),
}


class FormattingTab(BaseSettingsTab):
    """Formatting settings tab backed by typed draft state."""

    def __init__(self, view_model: SettingsViewModel, *, direction: Literal["input", "output"] = "output") -> None:
        self._direction: Literal["input", "output"] = direction
        self._vm = view_model
        self._extension_checks: dict[tuple[str, str], QCheckBox] = {}
        self._each_line_paragraph: QCheckBox = _cast(QCheckBox, None)
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
        self._mermaid_mode: QComboBox = _cast(QComboBox, None)
        self._dash_sep: QComboBox = _cast(QComboBox, None)
        self._asterisk_sep: QComboBox = _cast(QComboBox, None)
        self._underscore_sep: QComboBox = _cast(QComboBox, None)
        super().__init__()
        for name, key in (
            ("_page_break", "page_break_sep"),
            ("_section_break", "section_break_sep"),
            ("_horizontal_rule", "horizontal_rule_sep"),
            ("_bold_syntax", "bold_syntax"),
            ("_italic_syntax", "italic_syntax"),
            ("_strike_syntax", "strikethrough_syntax"),
            ("_highlight_syntax", "highlight_syntax"),
            ("_super_syntax", "superscript_syntax"),
            ("_sub_syntax", "subscript_syntax"),
            ("_ul_syntax", "unordered_list_syntax"),
            ("_indent_spaces", "indent_spaces"),
            ("_mermaid_mode", "mermaid_mode"),
            ("_dash_sep", "dash_sep"),
            ("_asterisk_sep", "asterisk_sep"),
            ("_underscore_sep", "underscore_sep"),
        ):
            combo = getattr(self, name)
            if combo is not None:
                self._wire_combo(combo, key)
        self._load_values()

    def _create_interface(self) -> None:
        if self._direction == "output":
            description = t("settings.markdown_extensions.description", "Configure generated Markdown.")
        else:
            description = t(
                "settings.markdown_extensions.input_description", "Configure Markdown recognition and DOCX output."
            )
        self.set_tab_description(description)
        self._create_extension_card(self._direction)
        if self._direction == "output":
            self._create_output_controls()
        else:
            self._create_input_controls()

    def _create_output_controls(self) -> None:
        # ── Shared Markdown output syntax ──────────────────────────────────────────
        _c3, f3 = self.add_settings_card(
            t("settings.formatting.syntax_selection", "Syntax Selection"),
            t(
                "settings.formatting.syntax_selection_tooltip",
                "These output styles currently apply primarily to ordinary DOCX-to-Markdown conversion. "
                "Other sources and specialized optimizers do not yet implement all choices. "
                "These choices do not restrict Markdown input recognition.",
            ),
            object_name="formattingOutputSyntaxCard",
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

        # ── DOCX → MD: Separator Mapping ───────────────────────────────
        _c2, f2 = self.add_settings_card(
            t("settings.formatting.separator_mapping", "Separator Mapping"),
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

    def _create_input_controls(self) -> None:
        # ── MD → DOCX: Separator Mapping ───────────────────────────────
        _c5, f5 = self.add_settings_card(
            t("settings.formatting.md_separator_mapping", "Separator Mapping"),
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

        # ── MD → Word: ordinary prose paragraph policy ────────────────
        _paragraph_card, paragraph_form = self.add_settings_card(
            t("settings.formatting.line_paragraph_section", "Paragraph handling"),
            object_name="formattingLineParagraphCard",
        )
        self._each_line_paragraph = self.create_settings_toggle(
            t("settings.formatting.each_line_paragraph", "Each line as a paragraph"),
            t(
                "settings.formatting.each_line_paragraph_help",
                "Create separate DOCX paragraphs for consecutive ordinary Markdown text lines. "
                "Explicit hard breaks, headings, lists, quotes, tables and code remain unchanged.",
            ),
        )
        self._each_line_paragraph.setObjectName("formattingEachLineParagraph")
        self.add_form_row(paragraph_form, "", self._each_line_paragraph)
        self._each_line_paragraph.toggled.connect(
            lambda checked: self._vm.set_field(SECTION_FORMATTING, "each_line_paragraph", checked)
        )

        # ── MD → DOCX: Mermaid ────────────────────────────────────────
        _c_mermaid, f_mermaid = self.add_settings_card(
            t("settings.formatting.mermaid_title", "Mermaid"),
            t(
                "settings.formatting.mermaid_tooltip",
                "Render Mermaid fenced blocks as images with the local Mermaid CLI, or keep the source as code. "
                "If image rendering is unavailable or fails, DocWen preserves the code block and reports a warning.",
            ),
            object_name="formattingMdMermaidCard",
        )
        self.add_form_description(
            f_mermaid,
            t("settings.formatting.mermaid_output_group", "Output method"),
        )
        self._mermaid_mode = self.create_combobox(
            _mermaid_options(),
            t(
                "settings.formatting.mermaid_mode_tooltip",
                "Choose how Mermaid fenced blocks are written to generated DOCX files.",
            ),
        )
        self._mermaid_mode.setObjectName("formattingMermaidMode")
        self.add_form_row(
            f_mermaid,
            t("settings.formatting.mermaid_mode_label", "Output:"),
            self._mermaid_mode,
        )
        self.add_form_description(
            f_mermaid,
            t("settings.formatting.mermaid_renderer_group", "Local renderer"),
        )
        self._mermaid_controls = MermaidControls(self, f_mermaid, self._vm)

    def _save_extensions(self, _checked: bool = False) -> None:
        values = {direction: dict(items) for direction, items in self._vm.config.formatting.markdown_extensions.items()}
        values[self._direction] = {
            name: self._extension_checks[(self._direction, name)].isChecked() for name in EXTENSION_NAMES
        }
        self._vm.set_field(SECTION_FORMATTING, "markdown_extensions", values)

    def _apply_extension_preset(self, direction: str) -> None:
        for (group, _name), checkbox in self._extension_checks.items():
            if group == direction:
                checkbox.blockSignals(True)
                checkbox.setChecked(True)
                checkbox.blockSignals(False)
        self._save_extensions()

    def _create_extension_card(self, direction: Literal["input", "output"]) -> None:
        _card, form = self.add_settings_card(
            t(f"settings.markdown_extensions.{direction}", direction.title()),
            t(f"settings.markdown_extensions.{direction}_hint", "Extensions are optional."),
            object_name=f"markdownExtensions{direction.title()}Card",
        )
        for name in EXTENSION_NAMES:
            label = t(f"settings.markdown_extensions.{name}", name.replace("_", " "))
            help_text = t(
                f"settings.markdown_extensions.{direction}_help.{name}",
                _EXTENSION_HELP_FALLBACKS[(direction, name)],
            )
            help_widget, checkbox = self._create_toggle_with_info(label, help_text)
            checkbox.setObjectName(f"markdownExtension{direction.title()}{name.title().replace('_', '')}")
            self._extension_checks[(direction, name)] = checkbox
            form.addRow(help_widget)
            checkbox.toggled.connect(self._save_extensions)
        preset = QPushButton(t("settings.markdown_extensions.obsidian_preset", "Use Obsidian extensions"))
        preset.setObjectName(f"markdownExtensions{direction.title()}Preset")
        preset.clicked.connect(lambda _checked=False, selected=direction: self._apply_extension_preset(selected))
        form.addRow(preset)

    def _wire_combo(self, combo: QComboBox, key: str) -> None:
        combo.currentIndexChanged.connect(
            lambda _idx, k=key, c=combo: self._vm.set_field(SECTION_FORMATTING, k, self.get_combo_data(c))
        )

    def _load_values(self) -> None:
        fmt = self._vm.config.formatting
        if self._direction == "input":
            self._each_line_paragraph.blockSignals(True)
            self._each_line_paragraph.setChecked(fmt.each_line_paragraph)
            self._each_line_paragraph.blockSignals(False)
        for (direction, name), checkbox in self._extension_checks.items():
            checkbox.blockSignals(True)
            checkbox.setChecked(fmt.markdown_extensions[direction][name])
            checkbox.blockSignals(False)
        if self._direction == "output":
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
        if self._direction == "input":
            self.set_combo_data(self._mermaid_mode, fmt.mermaid_mode)
            self._mermaid_controls.load(fmt.mermaid_cli_path)
            self.set_combo_data(self._dash_sep, fmt.dash_sep)
            self.set_combo_data(self._asterisk_sep, fmt.asterisk_sep)
            self.set_combo_data(self._underscore_sep, fmt.underscore_sep)

    def reload_from_config(self) -> None:
        self._load_values()


class MarkdownInputTab(FormattingTab):
    """Markdown input recognition and DOCX generation, with independent reset ownership."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        super().__init__(view_model, direction="input")
