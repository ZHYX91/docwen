"""Localized help must explain controls, preserve real line breaks and option labels."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

LOCALES = Path(__file__).resolve().parents[4] / "i18n" / "locales"
LOCALE_PATHS = sorted(LOCALES.glob("*.toml"))


@pytest.mark.unit
@pytest.mark.parametrize("path", LOCALE_PATHS, ids=lambda p: p.stem)
def test_date_help_is_explanatory_and_language_help_has_real_line_breaks(path: Path) -> None:
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    folder = data["settings"]["output"]["date_folder"]
    assert folder["format_tooltip"] != folder["format_label"]
    assert "\n" in folder["format_tooltip"]
    assert "\\n" not in data["settings"]["general"]["language_tooltip"]


@pytest.mark.unit
@pytest.mark.parametrize("path", LOCALE_PATHS, ids=lambda p: p.stem)
def test_table_header_help_separates_choices_and_uses_visible_labels(path: Path) -> None:
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    formatting = data["settings"]["formatting"]
    text = formatting["md_table_header_format_tooltip"]
    # Arabic uses a concise single sentence for the choices, followed by a separate recommendation.
    if path.stem == "ar_SA":
        assert "\n" in text
        return
    lines = text.splitlines()
    assert len(lines) == 4
    for line, key in zip(lines[:3], ["apply_format", "keep_markup", "clean_markup"], strict=True):
        assert line.startswith(formatting["options"][key])


@pytest.mark.unit
@pytest.mark.parametrize("locale", ["de_DE", "fr_FR", "pt_BR", "ru_RU"])
def test_previously_abbreviated_help_explains_all_options(locale: str) -> None:
    data = tomllib.loads((LOCALES / f"{locale}.toml").read_text(encoding="utf-8"))["settings"]
    formatting = data["formatting"]
    for key in ["md_body_format_tooltip", "md_heading_format_tooltip"]:
        lines = formatting[key].splitlines()
        assert len(lines) == 3
        for line, option in zip(lines, ["apply_format", "keep_markup", "clean_markup"], strict=True):
            assert line.startswith(formatting["options"][option])
    for key in ["embed_md_file_tooltip", "wiki_embed_image_tooltip"]:
        lines = data["link"][key].splitlines()[2:]
        assert len(lines) == 4
        for line, mode in zip(lines, ["keep", "extract_text", "remove", "embed"], strict=True):
            assert line.startswith(data["link"]["modes"][mode])


@pytest.mark.gui
def test_date_format_info_button_uses_help_instead_of_label(qapp) -> None:
    from PySide6.QtWidgets import QToolButton

    from docwen_gui.i18n import get_locale, set_locale, t
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.output_tab import OutputTab

    previous = get_locale()
    set_locale("en_US")
    page = OutputTab(SettingsViewModel())
    try:
        tips = [button.toolTip() for button in page.findChildren(QToolButton, "settingsInfoButton")]
        assert t("settings.output.date_folder.format_tooltip") in tips
        assert t("settings.output.date_folder.format_label") not in tips
    finally:
        page.close()
        set_locale(previous)


# Audited settings-info and About-tool help keys; repeated controls share a key.
INFO_HELP_KEYS = (
    "about.tools.easyofd",
    "about.tools.emoji",
    "about.tools.img2pdf",
    "about.tools.latex2mathml",
    "about.tools.lxml",
    "about.tools.numpy",
    "about.tools.olefile",
    "about.tools.onnxruntime",
    "about.tools.openpyxl",
    "about.tools.paddleocr",
    "about.tools.pandas",
    "about.tools.pdf2docx",
    "about.tools.pillow",
    "about.tools.pillow_heif",
    "about.tools.pymupdf",
    "about.tools.pymupdf4llm",
    "about.tools.pyside6",
    "about.tools.python_docx",
    "about.tools.pywin32",
    "about.tools.pyyaml",
    "about.tools.rapidocr",
    "about.tools.tomlkit",
    "settings.export.base64_compress_threshold_tooltip",
    "settings.export.base64_compress_tooltip",
    "settings.extraction.image_extraction_mode_tooltip",
    "settings.extraction.ocr_blockquote_title_enabled_tooltip",
    "settings.extraction.ocr_placement_mode_tooltip",
    "settings.formatting.asterisk_tooltip",
    "settings.formatting.body_format_tooltip",
    "settings.formatting.bold_syntax_tooltip",
    "settings.formatting.dash_tooltip",
    "settings.formatting.heading_format_tooltip",
    "settings.formatting.heading_merge_mode_tooltip",
    "settings.formatting.heading_merge_punctuation_tooltip",
    "settings.formatting.highlight_syntax_tooltip",
    "settings.formatting.horizontal_rule_tooltip",
    "settings.formatting.indent_spaces_tooltip",
    "settings.formatting.italic_syntax_tooltip",
    "settings.formatting.list_separator_tooltip",
    "settings.formatting.md_body_format_tooltip",
    "settings.formatting.md_heading_format_tooltip",
    "settings.formatting.md_table_header_format_tooltip",
    "settings.formatting.mermaid_mode_tooltip",
    "settings.formatting.page_break_tooltip",
    "settings.formatting.section_break_tooltip",
    "settings.formatting.strikethrough_syntax_tooltip",
    "settings.formatting.subscript_syntax_tooltip",
    "settings.formatting.superscript_syntax_tooltip",
    "settings.formatting.table_header_format_tooltip",
    "settings.formatting.table_style_tooltip",
    "settings.formatting.underscore_tooltip",
    "settings.formatting.unordered_list_tooltip",
    "settings.general.font_tooltip",
    "settings.general.language_tooltip",
    "settings.general.output_language_tooltip",
    "settings.general.scale_tooltip",
    "settings.general.theme_tooltip",
    "settings.general.window.default_mode_tooltip",
    "settings.gongwen.attachment_alignment_help",
    "settings.gongwen.fields_help",
    "settings.link.embed_md_file_tooltip",
    "settings.link.image_link_style_tooltip",
    "settings.link.markdown_embed_image_tooltip",
    "settings.link.markdown_link_tooltip",
    "settings.link.max_depth_tooltip",
    "settings.link.md_file_link_style_tooltip",
    "settings.link.wiki_embed_image_tooltip",
    "settings.link.wiki_link_tooltip",
    "settings.logging.console_colorize_tooltip",
    "settings.logging.console_format_tooltip",
    "settings.logging.console_level_tooltip",
    "settings.logging.directory_mode_tooltip",
    "settings.logging.enable_console_tooltip",
    "settings.logging.enable_logging_tooltip",
    "settings.logging.file_prefix_tooltip",
    "settings.logging.log_level_tooltip",
    "settings.logging.retention_days_tooltip",
    "settings.markdown_extensions.input_help.captions_references",
    "settings.markdown_extensions.input_help.extended_headings",
    "settings.markdown_extensions.input_help.structural_tables",
    "settings.markdown_extensions.input_help.typed_endnotes",
    "settings.markdown_extensions.output_help.captions_references",
    "settings.markdown_extensions.output_help.extended_headings",
    "settings.markdown_extensions.output_help.structural_tables",
    "settings.markdown_extensions.output_help.typed_endnotes",
    "settings.ocr.language_tooltip",
    "settings.ocr.recognize_tables_tooltip",
    "settings.output.date_folder.format_tooltip",
    "settings.output.output_mode_tooltip",
    "settings.spreadsheet.merge_mode_help",
    "settings.table_export.merge_strategy_help",
    "settings.text.numbering_comments_help",
    "settings.text.output_format_hint",
    "settings.text.render_mode_hint",
    "settings.text.scheme_tooltip",
)


@pytest.mark.unit
@pytest.mark.parametrize("path", LOCALE_PATHS, ids=lambda p: p.stem)
def test_all_audited_info_help_keys_are_present_and_have_no_literal_newlines(path: Path) -> None:
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    for key in INFO_HELP_KEYS:
        value = data
        for part in key.split("."):
            value = value[part]
        assert isinstance(value, str) and value.strip(), key
        assert "\\n" not in value, key
