"""Typed config dataclasses for Settings state.

Each dataclass maps to a logical config section.  SettingsViewModel
owns a single ``SettingsConfig`` instance and exposes per-section
subscription signals so individual tabs can observe changes without
tight coupling.

These dataclasses use ``__post_init__`` for value normalization but
remain plain-data — they carry no Qt dependency and can be shared
with the ApplicationController.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from docwen_core.text.heading_merge import DEFAULT_HEADING_MERGE_PUNCTUATION
from docwen_gui.window_behavior import DEFAULT_WINDOW_BEHAVIOR

# ── GUIConfig ──────────────────────────────────────────────────────────────


@dataclass
class GUIConfig:
    """General / Appearance settings."""

    language: str = "zh_CN"
    theme: str = "light"  # light / dark / system
    transparency_enabled: bool = False
    transparency_value: float = 1.0  # 0.20 .. 1.00
    remember_gui_state: bool = DEFAULT_WINDOW_BEHAVIOR.remember_gui_state
    auto_center: bool = DEFAULT_WINDOW_BEHAVIOR.auto_center
    expand_side_panels: bool = DEFAULT_WINDOW_BEHAVIOR.expand_side_panels
    default_mode: str = "single"  # single / batch
    md_default_template: str = "docx"  # docx / xlsx


# ── TextConfig ─────────────────────────────────────────────────────────────


@dataclass
class TextConfig:
    """Numbering / field-processor settings (Text tab)."""

    remove_numbering: bool = True
    add_numbering: bool = False
    default_scheme: str = "hierarchical_standard"
    numbering_schemes: dict = field(
        default_factory=lambda: {
            "settings": {"default_scheme": "hierarchical_standard", "order": []},
            "schemes": {},
        }
    )
    numbering_clean_rules: dict = field(
        default_factory=lambda: {
            "settings": {"order": []},
            "rules": [],
        }
    )
    field_processors: dict = field(
        default_factory=lambda: {
            "settings": {"order": []},
            "processors": {},
        }
    )
    heading_numbering_render_mode: str = "text"


# ── ProofreadConfig ────────────────────────────────────────────────────────


@dataclass
class ProofreadConfig:
    """Proofread engine and skip-rule toggles."""

    symbol_pairing: bool = True
    symbol_correction: bool = True
    typos_rule: bool = True
    sensitive_word: bool = True
    skip_code_blocks: bool = True
    skip_quote_blocks: bool = True
    symbol_mappings: dict = field(default_factory=dict)
    typos_dict: dict = field(default_factory=dict)
    sensitive_words: dict = field(default_factory=dict)


# ── ConversionDefaultsConfig ───────────────────────────────────────────────


@dataclass
class ConversionDefaultsConfig:
    """Per-category defaults used by conversion.

    Each field is a flat dict that maps config keys to their default
    values for document / spreadsheet / image / layout / other / export.
    """

    document: dict = field(default_factory=dict)
    spreadsheet: dict = field(default_factory=dict)
    image: dict = field(default_factory=dict)
    layout: dict = field(default_factory=dict)
    other: dict = field(default_factory=dict)
    export: dict = field(default_factory=dict)


# ── SoftwarePriorityConfig ─────────────────────────────────────────────────


@dataclass
class SoftwarePriorityConfig:
    """Priority lists for external office software per category."""

    word_processors: list[str] = field(default_factory=lambda: ["wps_writer", "msoffice_word", "libreoffice"])
    odt_conversion: list[str] = field(default_factory=lambda: ["msoffice_word", "libreoffice"])
    document_to_pdf: list[str] = field(default_factory=lambda: ["wps_writer", "msoffice_word", "libreoffice"])
    spreadsheet_processors: list[str] = field(
        default_factory=lambda: ["wps_spreadsheets", "msoffice_excel", "libreoffice"]
    )
    ods_conversion: list[str] = field(default_factory=lambda: ["msoffice_excel", "libreoffice"])
    spreadsheet_to_pdf: list[str] = field(default_factory=lambda: ["wps_spreadsheets", "msoffice_excel", "libreoffice"])
    pdf_to_office: list[str] = field(default_factory=lambda: ["msoffice_word", "libreoffice"])


# ── LinkConfig ─────────────────────────────────────────────────────────────


@dataclass
class LinkConfig:
    """Link processing settings."""

    image_link_style: str = "wiki_embed"
    md_file_link_style: str = "wiki_embed"
    wiki_link_mode: str = "hyperlink"
    markdown_link_mode: str = "hyperlink"
    wiki_embed_image_mode: str = "embed"
    markdown_embed_image_mode: str = "embed"
    embed_md_file_mode: str = "embed"
    max_depth: int = 3


# ── FormattingConfig ───────────────────────────────────────────────────────


@dataclass
class FormattingConfig:
    """DOCX<->MD formatting settings.

    Maps to the Formatting tab's combo-box and exact-text controls.
    """

    # DOCX -> MD: format processing
    body_format: str = "preserve"
    heading_format: str = "discard"
    table_header_format: str = "discard"

    # DOCX -> MD: separators
    page_break_sep: str = "---"
    section_break_sep: str = "***"
    horizontal_rule_sep: str = "___"

    # DOCX -> MD: syntax
    bold_syntax: str = "asterisk"
    italic_syntax: str = "asterisk"
    strikethrough_syntax: str = "extended"
    highlight_syntax: str = "extended"
    superscript_syntax: str = "html"
    subscript_syntax: str = "html"
    unordered_list_syntax: str = "dash"
    indent_spaces: int = 4

    # MD -> DOCX
    md_body_format: str = "apply"
    md_heading_format: str = "remove"
    md_table_header_format: str = "remove"
    heading_merge_mode: str = "punct_required"
    heading_merge_punctuation: str = DEFAULT_HEADING_MERGE_PUNCTUATION
    list_separator: str = "、"
    table_style_mode: str = "builtin"
    builtin_table_style: str = "three_line_table"
    custom_table_style_name: str = ""

    # MD -> DOCX separators
    dash_sep: str = "page_break"
    asterisk_sep: str = "section_break"
    underscore_sep: str = "horizontal_rule_1"


# ── OutputConfig ───────────────────────────────────────────────────────────


@dataclass
class OutputConfig:
    """Output directory and behavior settings."""

    output_mode: str = "source"  # source / custom
    custom_path: str = ""
    create_date_subfolder: bool = False
    date_folder_format: str = "%Y-%m-%d"
    auto_open_folder: bool = False
    save_intermediate_files: bool = False


# ── ExportConfig ───────────────────────────────────────────────────────────


@dataclass
class ExportConfig:
    """Export (image extraction / OCR placement / Base64) settings."""

    image_mode: str = "file"  # file / base64
    ocr_mode: str = "image_md"  # image_md / main_md
    ocr_title_enabled: bool = True
    ocr_title_text: str = ""
    base64_compress_enabled: bool = True
    base64_compress_threshold_kb: int = 100


# ── LoggingConfig ──────────────────────────────────────────────────────────


@dataclass
class LoggingConfig:
    """Logging system settings."""

    enable: bool = True
    level: str = "debug"
    file_prefix: str = "docwen"
    retention_days: int = 30
    console_enable: bool = True
    console_level: str = "info"
    console_format: str = ""
    console_colorize: str = "auto"
    directory_mode: str = "user"
    directory: str = ""


# ── SettingsConfig (aggregate root) ────────────────────────────────────────


@dataclass
class SettingsConfig:
    """Aggregate root for all settings state.

    Owned by SettingsViewModel.  Each section is a typed dataclass.
    The ``_dirty`` dict tracks which sections have unsaved changes.
    """

    gui: GUIConfig = field(default_factory=GUIConfig)
    text: TextConfig = field(default_factory=TextConfig)
    proofread: ProofreadConfig = field(default_factory=ProofreadConfig)
    conversion_defaults: ConversionDefaultsConfig = field(default_factory=ConversionDefaultsConfig)
    software_priority: SoftwarePriorityConfig = field(default_factory=SoftwarePriorityConfig)
    link: LinkConfig = field(default_factory=LinkConfig)
    formatting: FormattingConfig = field(default_factory=FormattingConfig)
    output: OutputConfig = field(default_factory=OutputConfig)
    export: ExportConfig = field(default_factory=ExportConfig)
    logging: LoggingConfig = field(default_factory=LoggingConfig)

    _dirty: set[str] = field(default_factory=set)

    def mark_dirty(self, section: str) -> None:
        self._dirty.add(section)

    def mark_clean(self, section: str | None = None) -> None:
        if section is None:
            self._dirty.clear()
        else:
            self._dirty.discard(section)

    @property
    def dirty_sections(self) -> frozenset[str]:
        return frozenset(self._dirty)

    @property
    def is_dirty(self) -> bool:
        return len(self._dirty) > 0
