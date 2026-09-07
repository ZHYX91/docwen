"""Content-format controls owned by incoming text and document settings."""

from __future__ import annotations

from contextlib import ExitStack

from PySide6.QtCore import QSignalBlocker
from PySide6.QtWidgets import QComboBox, QLineEdit

from ...i18n import t
from ...view_models.settings_vm import SECTION_DOCUMENT, SECTION_TEXT, SettingsViewModel
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


class DocumentContentControls:
    def __init__(self, tab: BaseSettingsTab, view_model: SettingsViewModel) -> None:
        self._tab = tab
        self._vm = view_model
        # ── DOCX → MD: Format Processing ────────────────────────────────
        _c1, f1 = self._tab.add_settings_card(
            f"{t('settings.formatting.docx_to_md_section', 'DOCX to MD')} — {t('settings.formatting.format_processing', 'Format Processing')}",
            t(
                "settings.formatting.format_processing_tooltip",
                "How to handle formatting when converting from DOCX to Markdown.",
            ),
            object_name="formattingDocxProcessingCard",
        )
        self._body_format = self._tab.create_combobox(
            _fmt_options(), t("settings.formatting.body_format_tooltip", "How to handle body text formatting")
        )
        self._tab.add_form_row(f1, t("settings.formatting.body_format_label", "Body Text Format:"), self._body_format)

        self._heading_format = self._tab.create_combobox(
            _fmt_options(), t("settings.formatting.heading_format_tooltip", "How to handle heading formatting")
        )
        self._tab.add_form_row(
            f1, t("settings.formatting.heading_format_label", "Heading Format:"), self._heading_format
        )

        self._table_header_format = self._tab.create_combobox(
            _fmt_options(),
            t("settings.formatting.table_header_format_tooltip", "How to handle table header formatting"),
        )
        self._tab.add_form_row(
            f1, t("settings.formatting.table_header_format_label", "Table Header Format:"), self._table_header_format
        )

        self._wire_combo(self._body_format, "body_format")
        self._wire_combo(self._heading_format, "heading_format")
        self._wire_combo(self._table_header_format, "table_header_format")

        self.reload_from_config()

    def _wire_combo(self, combo: QComboBox, key: str) -> None:
        combo.currentIndexChanged.connect(lambda _index: self._vm.set_field(SECTION_DOCUMENT, key, combo.currentData()))

    def reload_from_config(self) -> None:
        fmt = self._vm.config.document
        with ExitStack() as stack:
            for control in (
                self._body_format,
                self._heading_format,
                self._table_header_format,
            ):
                stack.enter_context(QSignalBlocker(control))
            self._tab.set_combo_data(self._body_format, fmt.body_format)
            self._tab.set_combo_data(self._heading_format, fmt.heading_format)
            self._tab.set_combo_data(self._table_header_format, fmt.table_header_format)


class MarkdownContentControls:
    def __init__(self, tab: BaseSettingsTab, view_model: SettingsViewModel) -> None:
        self._tab = tab
        self._vm = view_model
        # ── MD → DOCX: Format Processing ───────────────────────────────
        _c4, f4 = self._tab.add_settings_card(
            f"{t('settings.formatting.md_to_docx_section', 'MD to DOCX')} — {t('settings.formatting.md_format_processing', 'Format Processing')}",
            t(
                "settings.formatting.md_format_processing_tooltip",
                "How to handle formatting when converting from MD to DOCX.",
            ),
            object_name="formattingMdProcessingCard",
        )
        self._md_body_format = self._tab.create_combobox(
            _md_fmt_options(), t("settings.formatting.md_body_format_tooltip", "Body format markup handling")
        )
        self._tab.add_form_row(
            f4, t("settings.formatting.md_body_format_label", "Body Text Format:"), self._md_body_format
        )
        self._md_heading_format = self._tab.create_combobox(
            _md_fmt_options(), t("settings.formatting.md_heading_format_tooltip", "Heading format markup handling")
        )
        self._tab.add_form_row(
            f4, t("settings.formatting.md_heading_format_label", "Heading Format:"), self._md_heading_format
        )
        self._md_table_header_format = self._tab.create_combobox(
            _md_fmt_options(),
            t("settings.formatting.md_table_header_format_tooltip", "Table header format markup handling"),
        )
        self._tab.add_form_row(
            f4,
            t("settings.formatting.md_table_header_format_label", "Table Header Format:"),
            self._md_table_header_format,
        )
        self._heading_merge_mode = self._tab.create_combobox(
            _heading_merge_options(), t("settings.formatting.heading_merge_mode_tooltip", "Heading merge mode")
        )
        self._tab.add_form_row(
            f4, t("settings.formatting.heading_merge_mode_label", "Heading Merge Mode:"), self._heading_merge_mode
        )
        self._heading_merge_punctuation = QLineEdit(self._tab)
        self._heading_merge_punctuation.setToolTip(
            t(
                "settings.formatting.heading_merge_punctuation_tooltip",
                "Characters that trigger heading/body merging when punctuation is required. Empty disables punctuation-triggered merging.",
            )
        )
        self._tab.add_form_row(
            f4,
            t("settings.formatting.heading_merge_punctuation_label", "Merge punctuation:"),
            self._heading_merge_punctuation,
        )
        _template_card, template_form = self._tab.add_settings_card(
            t("settings.text.template_fill_section", "Template Filling"),
            t(
                "settings.text.template_fill_desc",
                "Join YAML list values when filling Word or spreadsheet templates. Spaces and an empty separator are preserved exactly.",
            ),
            object_name="textTemplateFillCard",
        )
        self._list_separator = QLineEdit(self._tab)
        self._list_separator.setToolTip(
            t(
                "settings.formatting.list_separator_tooltip",
                "Separator used when joining YAML list values into a string.",
            )
        )
        self._tab.add_form_row(
            template_form,
            t("settings.formatting.list_separator_label", "YAML List Separator") + ":",
            self._list_separator,
        )

        # ── MD → DOCX: Table Style ────────────────────────────────────
        _c_table, f_table = self._tab.add_settings_card(
            f"{t('settings.formatting.md_to_docx_section', 'MD to DOCX')} — {t('settings.formatting.table_style', 'Table Style')}",
            t(
                "settings.formatting.table_style_tooltip",
                "Choose the Word table style used for Markdown tables.",
            ),
            object_name="formattingMdTableStyleCard",
        )
        self._table_style_mode = self._tab.create_combobox(
            _table_style_mode_options(),
            t("settings.formatting.table_style_tooltip", "Use a built-in table style or a custom style name."),
        )
        self._tab.add_form_row(
            f_table,
            t("settings.formatting.table_style", "Table Style") + ":",
            self._table_style_mode,
        )
        self._builtin_table_style = self._tab.create_combobox(
            _builtin_table_style_options(),
            t("settings.formatting.table_style_tooltip", "Built-in table style for Markdown tables."),
        )
        self._tab.add_form_row(
            f_table,
            t("settings.formatting.builtin_style_radio", "Use built-in style") + ":",
            self._builtin_table_style,
        )
        self._custom_table_style_name = QLineEdit(self._tab)
        self._custom_table_style_name.setToolTip(
            t("settings.formatting.table_style_tooltip", "Custom Word table style name.")
        )
        self._tab.add_form_row(
            f_table,
            t("settings.formatting.custom_style_radio", "Use custom style name") + ":",
            self._custom_table_style_name,
        )

        self._wire_combo(self._md_body_format, "md_body_format")
        self._wire_combo(self._md_heading_format, "md_heading_format")
        self._wire_combo(self._md_table_header_format, "md_table_header_format")
        self._wire_combo(self._heading_merge_mode, "heading_merge_mode")
        self._heading_merge_mode.currentIndexChanged.connect(
            lambda _idx: self._sync_heading_merge_punctuation_control()
        )
        self._heading_merge_punctuation.textChanged.connect(
            lambda text: self._vm.set_field(SECTION_TEXT, "heading_merge_punctuation", text)
        )
        self._list_separator.textChanged.connect(lambda text: self._vm.set_field(SECTION_TEXT, "list_separator", text))
        self._wire_combo(self._table_style_mode, "table_style_mode")
        self._wire_combo(self._builtin_table_style, "builtin_table_style")
        self._table_style_mode.currentIndexChanged.connect(lambda _idx: self._sync_table_style_controls())
        self._custom_table_style_name.textChanged.connect(
            lambda text: self._vm.set_field(SECTION_TEXT, "custom_table_style_name", text.strip())
        )

        self.reload_from_config()

    def _wire_combo(self, combo: QComboBox, key: str) -> None:
        combo.currentIndexChanged.connect(lambda _index: self._vm.set_field(SECTION_TEXT, key, combo.currentData()))

    def reload_from_config(self) -> None:
        fmt = self._vm.config.text
        with ExitStack() as stack:
            for control in (
                self._builtin_table_style,
                self._custom_table_style_name,
                self._heading_merge_mode,
                self._heading_merge_punctuation,
                self._list_separator,
                self._md_body_format,
                self._md_heading_format,
                self._md_table_header_format,
                self._table_style_mode,
            ):
                stack.enter_context(QSignalBlocker(control))
            self._tab.set_combo_data(self._md_body_format, fmt.md_body_format)
            self._tab.set_combo_data(self._md_heading_format, fmt.md_heading_format)
            self._tab.set_combo_data(self._md_table_header_format, fmt.md_table_header_format)
            self._tab.set_combo_data(self._heading_merge_mode, fmt.heading_merge_mode)
            self._heading_merge_punctuation.setText(fmt.heading_merge_punctuation)
            self._sync_heading_merge_punctuation_control()
            self._list_separator.setText(fmt.list_separator)
            self._tab.set_combo_data(self._table_style_mode, fmt.table_style_mode)
            self._tab.set_combo_data(self._builtin_table_style, fmt.builtin_table_style)
            self._custom_table_style_name.setText(fmt.custom_table_style_name)
            self._sync_table_style_controls()

    def _sync_table_style_controls(self) -> None:
        use_builtin = self._tab.get_combo_data(self._table_style_mode) == "builtin"
        self._builtin_table_style.setEnabled(use_builtin)
        self._custom_table_style_name.setEnabled(not use_builtin)

    def _sync_heading_merge_punctuation_control(self) -> None:
        self._heading_merge_punctuation.setEnabled(
            self._tab.get_combo_data(self._heading_merge_mode) == "punct_required"
        )
