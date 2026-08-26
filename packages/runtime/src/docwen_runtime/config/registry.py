"""Single source of truth for config file layout.

Each :class:`ConfigFileSpec` declares a relative path within the config directory,
its dotted-key namespace for flat-format TOML routing, and zero or more GUI/CLI
groups that determine which configuration panel or command-line group owns it.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import PurePosixPath
from typing import Any


@dataclass(frozen=True)
class ConfigFileSpec:
    """Registry metadata for one physical config file.

    ``replace_sections`` lists top-level sections that are complete user-owned
    snapshots whenever they are present in the user file. Missing sections
    reveal shipped defaults; present sections replace them wholesale.

    ``keyed_list_sections`` lists top-level arrays of tables whose entries
    have stable ``id`` fields.  Shipped entries remain present across an
    upgrade, while matching user entries override them and user-only entries
    are appended.
    """

    rel_path: str
    namespace: tuple[str, ...]
    groups: tuple[str, ...] = ()
    replace_sections: frozenset[str] = frozenset()
    keyed_list_sections: frozenset[str] = frozenset()

    @property
    def posix_path(self) -> PurePosixPath:
        return PurePosixPath(self.rel_path)


@dataclass(frozen=True)
class ConfigGroupResetPlan:
    """Physical files and precise values owned by one settings reset group."""

    files: tuple[str, ...]
    dotted_keys: tuple[str, ...]


CONFIG_FILES: tuple[ConfigFileSpec, ...] = (
    ConfigFileSpec("gui.toml", ("gui",), ("general",)),
    ConfigFileSpec("output.toml", ("output",), ("output",)),
    ConfigFileSpec("logger.toml", ("logger",), ("logging",)),
    ConfigFileSpec("conversion.toml", ("conversion",), ("formatting",)),
    ConfigFileSpec("export.toml", ("export",), ("export", "conversion_defaults")),
    ConfigFileSpec("other.toml", ("other",), ("other", "conversion_defaults")),
    ConfigFileSpec("document.toml", ("document",), ("document", "conversion_defaults")),
    ConfigFileSpec("text.toml", ("text",), ("text", "conversion_defaults")),
    ConfigFileSpec("layout.toml", ("layout",), ("layout", "conversion_defaults")),
    ConfigFileSpec("spreadsheet.toml", ("spreadsheet",), ("spreadsheet", "conversion_defaults")),
    ConfigFileSpec("image.toml", ("image",), ("image", "conversion_defaults")),
    ConfigFileSpec("link.toml", ("link",), ("link",)),
    ConfigFileSpec("software.toml", ("software",), ("software_priority", "software")),
    ConfigFileSpec("optimize.toml", ("optimize",), ()),
    ConfigFileSpec("field_processors.toml", ("field_processors",), ("text",)),
    ConfigFileSpec("numbering/add.toml", ("numbering", "add"), ("text",)),
    ConfigFileSpec(
        "numbering/cleanup.toml",
        ("numbering", "cleanup"),
        ("text",),
        keyed_list_sections=frozenset({"rules"}),
    ),
    ConfigFileSpec("proofread/engine.toml", ("proofread", "engine"), ("proofread",)),
    ConfigFileSpec("proofread/skip.toml", ("proofread", "skip"), ("proofread",)),
    ConfigFileSpec("proofread/pairs.toml", ("proofread", "pairs"), ("proofread",)),
    ConfigFileSpec(
        "proofread/symbol_map.toml",
        ("proofread", "symbol_map"),
        ("proofread",),
        replace_sections=frozenset({"entries"}),
    ),
    ConfigFileSpec(
        "proofread/typos.toml",
        ("proofread", "typos"),
        ("proofread",),
        replace_sections=frozenset({"entries"}),
    ),
    ConfigFileSpec(
        "proofread/sensitive_words.toml",
        ("proofread", "sensitive_words"),
        ("proofread",),
        replace_sections=frozenset({"entries"}),
    ),
)

_BY_PATH: dict[str, ConfigFileSpec] = {spec.rel_path: spec for spec in CONFIG_FILES}


# Most settings groups own complete files and need no override here.  These
# groups cross physical-file boundaries or share a file with another tab, so
# their logical ownership must be expressed as dotted values.  This registry
# remains the single source consumed by both GUI and CLI reset entry points.
_GROUP_RESET_DOTTED_KEYS: dict[str, tuple[str, ...]] = {
    "general": (
        "gui.theme",
        "gui.window",
        "gui.transparency",
        "gui.language",
    ),
    "text": (
        "text.remove_numbering",
        "text.add_numbering",
        "text.numbering_scheme",
        "text.heading_numbering_render_mode",
        "gui.template.md_default_template",
        "numbering.add.settings.default_scheme",
    ),
    "proofread": (
        "proofread.engine.enable_symbol_pairing",
        "proofread.engine.enable_symbol_correction",
        "proofread.engine.enable_typos_rule",
        "proofread.engine.enable_sensitive_word",
        "proofread.skip.code_blocks",
        "proofread.skip.quote_blocks",
    ),
    "export": (
        "conversion.ocr_output.show_blockquote_title",
        "conversion.ocr_output.blockquote_title_override_by_locale",
        "conversion.export.base64_compress_enabled",
        "conversion.export.base64_compress_threshold_kb",
    ),
    "formatting": (
        "conversion.docx_to_md.preserve_formatting",
        "conversion.docx_to_md.preserve_heading_formatting",
        "conversion.docx_to_md.preserve_table_header_formatting",
        "conversion.md_to_docx.formatting_mode",
        "conversion.md_to_docx.heading_formatting_mode",
        "conversion.md_to_docx.table_header_formatting_mode",
        "conversion.md_to_docx.heading_merge_mode",
        "conversion.md_to_docx.heading_merge_punctuation",
        "conversion.md_to_docx.list_separator",
        "conversion.syntax.bold",
        "conversion.syntax.italic",
        "conversion.syntax.strikethrough",
        "conversion.syntax.highlight",
        "conversion.syntax.superscript",
        "conversion.syntax.subscript",
        "conversion.syntax.unordered_list",
        "conversion.syntax.indent_spaces",
        "conversion.horizontal_rule.docx_to_md.page_break",
        "conversion.horizontal_rule.docx_to_md.section_break",
        "conversion.horizontal_rule.docx_to_md.horizontal_rule",
        "conversion.horizontal_rule.md_to_docx.dash",
        "conversion.horizontal_rule.md_to_docx.asterisk",
        "conversion.horizontal_rule.md_to_docx.underscore",
        "document.style.table.md_to_docx.table_style_mode",
        "document.style.table.md_to_docx.builtin_style_key",
        "document.style.table.md_to_docx.custom_style_name",
    ),
    "document": (
        "document.to_md_keep_images",
        "document.to_md_enable_ocr",
        "document.to_md_table_merge_export_strategy",
        "document.to_md_remove_numbering",
        "document.to_md_add_numbering",
        "document.to_md_default_scheme",
        "document.to_md_enable_optimization",
        "document.to_md_optimization_type",
        "software.default_priority.word_processors",
        "software.special_conversions.odt",
        "software.special_conversions.document_to_pdf",
    ),
    "spreadsheet": (
        "spreadsheet.to_md_keep_images",
        "spreadsheet.to_md_enable_ocr",
        "spreadsheet.to_md_table_merge_export_strategy",
        "spreadsheet.merge_mode",
        "software.default_priority.spreadsheet_processors",
        "software.special_conversions.ods",
        "software.special_conversions.spreadsheet_to_pdf",
    ),
    "layout": (
        "layout.to_md_keep_images",
        "layout.to_md_enable_ocr",
        "layout.to_md_enable_optimization",
        "layout.to_md_optimization_type",
        "layout.render_dpi",
        "software.special_conversions.pdf_to_office",
    ),
    "link": (
        "link.format.image_link_style",
        "link.format.md_file_link_style",
        "link.non_embed_links.wiki_mode",
        "link.non_embed_links.markdown_mode",
        "link.embed_links.wiki_image_mode",
        "link.embed_links.markdown_image_mode",
        "link.embed_links.md_file_mode",
        "link.embedding.max_depth",
    ),
    "other": (
        "other.to_md_keep_images",
        "other.to_md_enable_ocr",
    ),
    "output": (
        "output.intermediate_files.save_to_output",
        "output.directory.mode",
        "output.directory.custom_path",
        "output.directory.create_date_subfolder",
        "output.directory.date_folder_format",
        "output.behavior.auto_open_folder",
    ),
    "logging": (
        "logger.enable",
        "logger.level",
        "logger.file_prefix",
        "logger.retention_days",
        "logger.console_enable",
        "logger.console_level",
        "logger.console_format",
        "logger.console_colorize",
        "logger.directory_mode",
        "logger.directory",
    ),
}


# A registry group describes panel ownership, which is intentionally broader
# than the destructive scope of "restore this page".  User-authored editors
# remain visible from their panel but are preserved by a current-page reset.
_GROUP_RESET_PRESERVED_FILES: dict[str, frozenset[str]] = {
    "text": frozenset({"field_processors.toml", "numbering/cleanup.toml"}),
    "proofread": frozenset(
        {
            "proofread/pairs.toml",
            "proofread/symbol_map.toml",
            "proofread/typos.toml",
            "proofread/sensitive_words.toml",
        }
    ),
}


def all_specs() -> tuple[ConfigFileSpec, ...]:
    return CONFIG_FILES


def get_spec(rel_path: str) -> ConfigFileSpec | None:
    normalized = rel_path.replace("\\", "/")
    return _BY_PATH.get(normalized)


def require_spec(rel_path: str) -> ConfigFileSpec:
    spec = get_spec(rel_path)
    if spec is None:
        raise KeyError(rel_path)
    return spec


def specs_for_group(group: str) -> tuple[ConfigFileSpec, ...]:
    return tuple(spec for spec in CONFIG_FILES if group in spec.groups)


def spec_for_dotted_key(dotted_key: str) -> ConfigFileSpec | None:
    parts = tuple(part for part in dotted_key.split(".") if part)
    matches = [spec for spec in CONFIG_FILES if parts[: len(spec.namespace)] == spec.namespace]
    if not matches:
        return None
    return max(matches, key=lambda spec: len(spec.namespace))


def relative_key_for_spec(spec: ConfigFileSpec, dotted_key: str) -> tuple[str, ...]:
    parts = tuple(part for part in dotted_key.split(".") if part)
    if parts[: len(spec.namespace)] != spec.namespace:
        raise KeyError(dotted_key)
    return parts[len(spec.namespace) :]


def reset_plan_for_group(group: str) -> ConfigGroupResetPlan:
    """Return the logical reset plan for a registry group.

    A dotted-key override makes the corresponding group-owned file partial,
    while keys routed to other files are additional cross-file ownership.
    """

    specs = specs_for_group(group)
    dotted_keys = _GROUP_RESET_DOTTED_KEYS.get(group, ())
    group_paths = {spec.rel_path for spec in specs}
    preserved_paths = _GROUP_RESET_PRESERVED_FILES.get(group, frozenset())
    unknown_preserved_paths = preserved_paths - group_paths
    if unknown_preserved_paths:
        unknown = ", ".join(sorted(unknown_preserved_paths))
        raise RuntimeError(f"Reset-preserved files are not owned by {group!r}: {unknown}")
    partial_group_paths: set[str] = set()
    for dotted_key in dotted_keys:
        spec = spec_for_dotted_key(dotted_key)
        if spec is None:
            raise RuntimeError(f"Reset key is not registered: {dotted_key}")
        if spec.rel_path in group_paths:
            partial_group_paths.add(spec.rel_path)

    files = tuple(
        spec.rel_path
        for spec in specs
        if spec.rel_path not in partial_group_paths and spec.rel_path not in preserved_paths
    )
    return ConfigGroupResetPlan(files=files, dotted_keys=dotted_keys)


def wrap_namespace(data: dict[str, Any], namespace: tuple[str, ...]) -> dict[str, Any]:
    wrapped: dict[str, Any] = data
    for key in reversed(namespace):
        wrapped = {key: wrapped}
    return wrapped


def unwrap_namespace(data: dict[str, Any], namespace: tuple[str, ...]) -> dict[str, Any]:
    current: Any = data
    for key in namespace:
        if not isinstance(current, dict):
            return {}
        current = current.get(key, {})
    return current if isinstance(current, dict) else {}
