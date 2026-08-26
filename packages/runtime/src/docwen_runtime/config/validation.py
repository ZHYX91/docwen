"""Semantic admission for trusted runtime configuration snapshots.

The configuration registry intentionally allows sparse user overrides and
unknown extension keys. Known shipped keys must retain their declared value
shape and exact enum/range values before a snapshot is trusted.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, MutableMapping
from copy import deepcopy
from typing import Any


class ConfigSemanticError(ValueError):
    """Raised when a parsed configuration value violates its runtime contract."""


def _describe_type(value: object) -> str:
    return type(value).__name__


def _validate_known_shapes(value: Any, shipped: Any, *, path: str) -> None:
    """Validate known keys against the shipped tree without rejecting extensions."""
    if isinstance(shipped, Mapping):
        if not isinstance(value, Mapping):
            raise ConfigSemanticError(f"{path} must be a table, got {_describe_type(value)}")
        for key, shipped_value in shipped.items():
            if key in value:
                _validate_known_shapes(value[key], shipped_value, path=f"{path}.{key}")
        return

    if isinstance(shipped, bool):
        if not isinstance(value, bool):
            raise ConfigSemanticError(f"{path} must be bool, got {_describe_type(value)}")
        return

    if isinstance(shipped, int):
        if not isinstance(value, int) or isinstance(value, bool):
            raise ConfigSemanticError(f"{path} must be int, got {_describe_type(value)}")
        return

    if isinstance(shipped, float):
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            raise ConfigSemanticError(f"{path} must be numeric, got {_describe_type(value)}")
        return

    if isinstance(shipped, str):
        if not isinstance(value, str):
            raise ConfigSemanticError(f"{path} must be str, got {_describe_type(value)}")
        return

    if isinstance(shipped, list) and not isinstance(value, list):
        raise ConfigSemanticError(f"{path} must be list, got {_describe_type(value)}")


def _read_path(data: Mapping[str, Any], dotted: str) -> Any:
    current: Any = data
    for part in dotted.split("."):
        if not isinstance(current, Mapping) or part not in current:
            return None
        current = current[part]
    return current


def _write_path(data: MutableMapping[str, Any], dotted: str, value: Any) -> None:
    parts = dotted.split(".")
    current = data
    for part in parts[:-1]:
        child = current.get(part)
        if not isinstance(child, MutableMapping):
            return
        current = child
    if parts[-1] in current:
        current[parts[-1]] = value


def _canonical_choice(
    data: dict[str, Any],
    dotted: str,
    choices: frozenset[str],
    *,
    label: str,
) -> None:
    value = _read_path(data, dotted)
    if value is None:
        return
    if not isinstance(value, str):
        raise ConfigSemanticError(f"{label} must be a string")
    if value not in choices:
        allowed = "/".join(sorted(choices))
        raise ConfigSemanticError(f"{label} must be one of {allowed}")


def _mapping_at(data: Mapping[str, Any], dotted: str, *, label: str) -> Mapping[str, Any] | None:
    value = _read_path(data, dotted)
    if value is None:
        return None
    if not isinstance(value, Mapping):
        raise ConfigSemanticError(f"{label} must be a table")
    return value


def _validate_string_list_value(
    value: object,
    *,
    label: str,
    allow_empty_items: bool = False,
    allow_empty_list: bool = True,
    unique: bool = False,
    strip_items: bool = False,
    exact_length: int | None = None,
) -> None:
    if not isinstance(value, list):
        raise ConfigSemanticError(f"{label} must be a list")
    if not allow_empty_list and not value:
        raise ConfigSemanticError(f"{label} must contain at least one string")
    if exact_length is not None and len(value) != exact_length:
        raise ConfigSemanticError(f"{label} must contain exactly {exact_length} strings")

    seen: set[str] = set()
    for index, item in enumerate(value):
        if not isinstance(item, str):
            raise ConfigSemanticError(f"{label}[{index}] must be a string")
        if not allow_empty_items and not item.strip():
            raise ConfigSemanticError(f"{label}[{index}] must be non-empty")
        if strip_items:
            value[index] = item.strip()
            item = value[index]
        if unique:
            identity = item.strip()
            if identity in seen:
                raise ConfigSemanticError(f"{label} must not contain duplicate strings")
            seen.add(identity)


def _validate_string_list_at(
    data: Mapping[str, Any],
    dotted: str,
    *,
    label: str,
    allow_empty_items: bool = False,
    allow_empty_list: bool = True,
    unique: bool = False,
    strip_items: bool = False,
) -> None:
    value = _read_path(data, dotted)
    if value is None:
        return
    _validate_string_list_value(
        value,
        label=label,
        allow_empty_items=allow_empty_items,
        allow_empty_list=allow_empty_list,
        unique=unique,
        strip_items=strip_items,
    )


def _validate_optional_string(
    record: Mapping[str, Any],
    key: str,
    *,
    label: str,
    non_empty: bool = False,
) -> None:
    if key not in record:
        return
    value = record[key]
    if not isinstance(value, str) or (non_empty and not value.strip()):
        qualifier = "a non-empty string" if non_empty else "a string"
        raise ConfigSemanticError(f"{label}.{key} must be {qualifier}")


def _validate_optional_bool(record: Mapping[str, Any], key: str, *, label: str) -> None:
    if key in record and not isinstance(record[key], bool):
        raise ConfigSemanticError(f"{label}.{key} must be bool")


def _validate_dynamic_record_map(
    data: Mapping[str, Any],
    dotted: str,
    *,
    label: str,
) -> Mapping[str, Any] | None:
    records = _mapping_at(data, dotted, label=label)
    if records is None:
        return None
    for record_id, record in records.items():
        if not isinstance(record_id, str) or not record_id.strip():
            raise ConfigSemanticError(f"{label} keys must be non-empty strings")
        if record_id != record_id.strip():
            raise ConfigSemanticError(f"{label} keys must not have surrounding whitespace")
        if not isinstance(record, Mapping):
            raise ConfigSemanticError(f"{label}.{record_id} must be a table")
    return records


def _validate_logger(data: dict[str, Any]) -> None:
    levels = frozenset({"debug", "info", "warning", "error", "critical"})
    _canonical_choice(data, "level", levels, label="logger.level")
    _canonical_choice(data, "console_level", levels, label="logger.console_level")
    _canonical_choice(
        data,
        "console_colorize",
        frozenset({"auto", "always", "never"}),
        label="logger.console_colorize",
    )
    _canonical_choice(
        data,
        "directory_mode",
        frozenset({"user", "temp", "custom"}),
        label="logger.directory_mode",
    )

    retention_days = data.get("retention_days")
    if retention_days is not None and (
        not isinstance(retention_days, int) or isinstance(retention_days, bool) or retention_days <= 0
    ):
        raise ConfigSemanticError("logger.retention_days must be a positive integer")

    file_prefix = data.get("file_prefix")
    if file_prefix is not None:
        if not isinstance(file_prefix, str) or not file_prefix.strip():
            raise ConfigSemanticError("logger.file_prefix must be a non-empty string")
        normalized_prefix = file_prefix.strip()
        if re.search(r'[\\/*?:"<>|]', normalized_prefix):
            raise ConfigSemanticError("logger.file_prefix contains invalid filename characters")
        data["file_prefix"] = normalized_prefix

    if data.get("directory_mode") == "custom":
        directory = data.get("directory")
        if not isinstance(directory, str) or not directory.strip():
            raise ConfigSemanticError("logger.directory is required for custom directory mode")
        data["directory"] = directory.strip()


def _validate_output(data: dict[str, Any]) -> None:
    _canonical_choice(
        data,
        "directory.mode",
        frozenset({"source", "custom"}),
        label="output.directory.mode",
    )


def _validate_gui(data: dict[str, Any]) -> None:
    _canonical_choice(
        data,
        "theme.default_theme",
        frozenset({"light", "dark", "system"}),
        label="gui.theme.default_theme",
    )
    locale = _read_path(data, "language.locale")
    if locale is not None:
        if not isinstance(locale, str) or not locale.strip():
            raise ConfigSemanticError("gui.language.locale must be a non-empty string")
        _write_path(data, "language.locale", locale.strip())
    _validate_string_list_at(
        data,
        "history.recent_files",
        label="gui.history.recent_files",
        allow_empty_items=True,
    )
    _validate_string_list_at(
        data,
        "history.recent_templates",
        label="gui.history.recent_templates",
        allow_empty_items=True,
    )


def _validate_conversion(data: dict[str, Any]) -> None:
    threshold = _read_path(data, "export.base64_compress_threshold_kb")
    if threshold is not None and (not isinstance(threshold, int) or isinstance(threshold, bool) or threshold <= 0):
        raise ConfigSemanticError("conversion.export.base64_compress_threshold_kb must be a positive integer")

    overrides = _mapping_at(
        data,
        "ocr_output.blockquote_title_override_by_locale",
        label="conversion.ocr_output.blockquote_title_override_by_locale",
    )
    if overrides is not None:
        for locale, title in overrides.items():
            if not isinstance(locale, str) or not locale.strip():
                raise ConfigSemanticError(
                    "conversion.ocr_output.blockquote_title_override_by_locale keys must be non-empty strings"
                )
            if not isinstance(title, str):
                raise ConfigSemanticError(
                    f"conversion.ocr_output.blockquote_title_override_by_locale.{locale} must be a string"
                )


def _validate_export(data: dict[str, Any]) -> None:
    _canonical_choice(
        data,
        "to_md_image_extraction_mode",
        frozenset({"file", "base64", "embed", "omit"}),
        label="export.to_md_image_extraction_mode",
    )
    _canonical_choice(
        data,
        "to_md_ocr_placement_mode",
        frozenset({"image_md", "main_md"}),
        label="export.to_md_ocr_placement_mode",
    )


def _validate_link(data: dict[str, Any]) -> None:
    _canonical_choice(
        data,
        "format.image_link_style",
        frozenset({"markdown_embed", "markdown_link", "wiki_embed", "wiki_link"}),
        label="link.format.image_link_style",
    )
    _canonical_choice(
        data,
        "format.md_file_link_style",
        frozenset({"markdown_link", "wiki_embed", "wiki_link"}),
        label="link.format.md_file_link_style",
    )
    non_embed_choices = frozenset({"extract_text", "hyperlink", "keep", "remove"})
    _canonical_choice(
        data,
        "non_embed_links.wiki_mode",
        non_embed_choices,
        label="link.non_embed_links.wiki_mode",
    )
    _canonical_choice(
        data,
        "non_embed_links.markdown_mode",
        non_embed_choices,
        label="link.non_embed_links.markdown_mode",
    )
    embed_choices = frozenset({"embed", "extract_text", "keep", "remove"})
    for key in ("wiki_image_mode", "markdown_image_mode", "md_file_mode"):
        _canonical_choice(
            data,
            f"embed_links.{key}",
            embed_choices,
            label=f"link.embed_links.{key}",
        )

    max_depth = _read_path(data, "embedding.max_depth")
    if max_depth is not None and (
        not isinstance(max_depth, int) or isinstance(max_depth, bool) or not 1 <= max_depth <= 20
    ):
        raise ConfigSemanticError("link.embedding.max_depth must be an integer from 1 to 20")

    _validate_string_list_at(
        data,
        "path_resolution.search_dirs",
        label="link.path_resolution.search_dirs",
    )
    error_choices = frozenset({"ignore", "keep", "placeholder"})
    for key in ("file_not_found", "circular_reference", "max_depth_reached"):
        _canonical_choice(
            data,
            f"error_handling.{key}",
            error_choices,
            label=f"link.error_handling.{key}",
        )


def _validate_software(data: dict[str, Any]) -> None:
    for section_name in ("default_priority", "special_conversions"):
        section = _mapping_at(data, section_name, label=f"software.{section_name}")
        if section is None:
            continue
        for priority_name, value in section.items():
            if not isinstance(priority_name, str) or not priority_name.strip():
                raise ConfigSemanticError(f"software.{section_name} keys must be non-empty strings")
            if priority_name != priority_name.strip():
                raise ConfigSemanticError(f"software.{section_name} keys must not have surrounding whitespace")
            _validate_string_list_value(
                value,
                label=f"software.{section_name}.{priority_name}",
                unique=True,
                strip_items=True,
            )


def _validate_optimize(data: dict[str, Any]) -> None:
    _validate_string_list_at(
        data,
        "settings.order",
        label="optimize.settings.order",
        unique=True,
        strip_items=True,
    )
    records = _validate_dynamic_record_map(data, "types", label="optimize.types")
    if records is None:
        return
    for type_id, record in records.items():
        _validate_optional_bool(record, "enabled", label=f"optimize.types.{type_id}")


def _validate_field_processors(data: dict[str, Any]) -> None:
    _validate_string_list_at(
        data,
        "settings.order",
        label="field_processors.settings.order",
        unique=True,
        strip_items=True,
    )
    records = _validate_dynamic_record_map(data, "processors", label="field_processors.processors")
    if records is None:
        return
    for processor_id, record in records.items():
        label = f"field_processors.processors.{processor_id}"
        _validate_optional_string(record, "module", label=label, non_empty=True)
        if "module" not in record:
            raise ConfigSemanticError(f"{label}.module must be a non-empty string")
        if isinstance(record, MutableMapping):
            record["module"] = record["module"].strip()
        for key in ("name", "name_key", "description", "description_key"):
            _validate_optional_string(record, key, label=label)
        for key in ("enabled", "is_system"):
            _validate_optional_bool(record, key, label=label)
        if "locales" in record:
            _validate_string_list_value(
                record["locales"],
                label=f"{label}.locales",
                allow_empty_list=False,
                unique=True,
                strip_items=True,
            )


def _validate_numbering_add(data: dict[str, Any]) -> None:
    _validate_string_list_at(
        data,
        "settings.order",
        label="numbering.add.settings.order",
        unique=True,
        strip_items=True,
    )
    settings = _mapping_at(data, "settings", label="numbering.add.settings")
    if settings is not None:
        _validate_optional_string(
            settings,
            "default_scheme",
            label="numbering.add.settings",
            non_empty=True,
        )
        if "default_scheme" in settings and isinstance(settings, MutableMapping):
            settings["default_scheme"] = settings["default_scheme"].strip()

    number_styles = _validate_dynamic_record_map(data, "number_styles", label="numbering.add.number_styles")
    if number_styles is not None:
        for style_id, record in number_styles.items():
            label = f"numbering.add.number_styles.{style_id}"
            for key in ("name", "description"):
                _validate_optional_string(record, key, label=label)

    schemes = _validate_dynamic_record_map(data, "schemes", label="numbering.add.schemes")
    if schemes is None:
        return
    for scheme_id, record in schemes.items():
        label = f"numbering.add.schemes.{scheme_id}"
        for key in ("name", "name_key", "description", "description_key"):
            _validate_optional_string(record, key, label=label)
        for key in ("enabled", "is_system"):
            _validate_optional_bool(record, key, label=label)
        if "locales" in record:
            _validate_string_list_value(
                record["locales"],
                label=f"{label}.locales",
                allow_empty_list=False,
                unique=True,
                strip_items=True,
            )
        has_usable_level = False
        for level in range(1, 10):
            key = f"level_{level}"
            if key not in record:
                continue
            level_data = record[key]
            if not isinstance(level_data, Mapping):
                raise ConfigSemanticError(f"{label}.{key} must be a table")
            _validate_optional_string(level_data, "format", label=f"{label}.{key}")
            if "format" not in level_data:
                raise ConfigSemanticError(f"{label}.{key}.format must be a string")
            format_text = level_data["format"]
            if format_text.strip():
                has_usable_level = True
            for reference_level, style in re.findall(r"\{(\d+)\.(\w+)\}", format_text):
                if not 1 <= int(reference_level) <= 9:
                    raise ConfigSemanticError(f"{label}.{key} contains an out-of-range level reference")
                if style not in {
                    "arabic_circled",
                    "arabic_full",
                    "arabic_half",
                    "chinese_lower",
                    "chinese_upper",
                    "letter_lower",
                    "letter_upper",
                    "roman_lower",
                    "roman_upper",
                }:
                    raise ConfigSemanticError(f"{label}.{key} contains an unsupported numbering style")
        if record.get("enabled", True) is True and not has_usable_level:
            raise ConfigSemanticError(f"{label} must define at least one non-empty level format when enabled")


def _validate_numbering_cleanup(data: dict[str, Any]) -> None:
    _validate_string_list_at(
        data,
        "settings.order",
        label="numbering.cleanup.settings.order",
        unique=True,
        strip_items=True,
    )
    rules = data.get("rules")
    if rules is None:
        return
    if not isinstance(rules, list):
        raise ConfigSemanticError("numbering.cleanup.rules must be a list")

    seen_ids: set[str] = set()
    for index, rule in enumerate(rules):
        label = f"numbering.cleanup.rules[{index}]"
        if not isinstance(rule, Mapping):
            raise ConfigSemanticError(f"{label} must be a table")
        rule_id = rule.get("id")
        if not isinstance(rule_id, str) or not rule_id.strip():
            raise ConfigSemanticError(f"{label}.id must be a non-empty string")
        if rule_id != rule_id.strip():
            raise ConfigSemanticError(f"{label}.id must not have surrounding whitespace")
        if rule_id in seen_ids:
            raise ConfigSemanticError("numbering.cleanup.rules must not contain duplicate ids")
        seen_ids.add(rule_id)
        for key in ("name", "name_key", "description", "description_key"):
            _validate_optional_string(rule, key, label=label)
        for key in ("enabled", "is_system"):
            _validate_optional_bool(rule, key, label=label)
        pattern = rule.get("pattern")
        if not isinstance(pattern, str) or not pattern:
            raise ConfigSemanticError(f"{label}.pattern must be a non-empty string")
        try:
            re.compile(pattern)
        except re.error as exc:
            raise ConfigSemanticError(f"{label}.pattern must be a valid regular expression") from exc
        level = rule.get("level")
        if not isinstance(level, int) or isinstance(level, bool) or not 1 <= level <= 5:
            raise ConfigSemanticError(f"{label}.level must be an integer from 1 to 5")


def _validate_proofread_pairs(data: dict[str, Any]) -> None:
    items = data.get("items")
    if items is None:
        return
    if not isinstance(items, list):
        raise ConfigSemanticError("proofread.pairs.items must be a list")
    seen_pairs: set[tuple[str, str]] = set()
    closing_to_opening: dict[str, str] = {}
    for index, pair in enumerate(items):
        if isinstance(pair, Mapping):
            has_source_target = "source" in pair or "target" in pair
            has_open_close = "open" in pair or "close" in pair
            if has_source_target and has_open_close:
                raise ConfigSemanticError(f"proofread.pairs.items[{index}] cannot mix source/target with open/close")
            if has_source_target:
                if "source" not in pair or "target" not in pair:
                    raise ConfigSemanticError(f"proofread.pairs.items[{index}] must contain both source and target")
                pair = [pair["source"], pair["target"]]
            elif has_open_close:
                if "open" not in pair or "close" not in pair:
                    raise ConfigSemanticError(f"proofread.pairs.items[{index}] must contain both open and close")
                pair = [pair["open"], pair["close"]]
            else:
                raise ConfigSemanticError(
                    f"proofread.pairs.items[{index}] must be a pair or a source/target or open/close table"
                )
            items[index] = pair
        _validate_string_list_value(
            pair,
            label=f"proofread.pairs.items[{index}]",
            exact_length=2,
        )
        identity = (pair[0], pair[1])
        if identity in seen_pairs:
            raise ConfigSemanticError("proofread.pairs.items must not contain duplicate pairs")
        seen_pairs.add(identity)
        previous_opening = closing_to_opening.setdefault(pair[1], pair[0])
        if previous_opening != pair[0]:
            raise ConfigSemanticError("proofread.pairs.items cannot map one closing symbol to multiple openings")


def _validate_proofread_entries(data: dict[str, Any], *, namespace: str) -> None:
    entries = _mapping_at(data, "entries", label=f"proofread.{namespace}.entries")
    if entries is None:
        return
    for key, value in entries.items():
        if not isinstance(key, str) or not key.strip():
            raise ConfigSemanticError(f"proofread.{namespace}.entries keys must be non-empty strings")
        _validate_string_list_value(
            value,
            label=f"proofread.{namespace}.entries.{key}",
        )


def _validate_proofread_symbol_map(data: dict[str, Any]) -> None:
    _validate_proofread_entries(data, namespace="symbol_map")


def _validate_proofread_typos(data: dict[str, Any]) -> None:
    _validate_proofread_entries(data, namespace="typos")


def _validate_proofread_sensitive_words(data: dict[str, Any]) -> None:
    _validate_proofread_entries(data, namespace="sensitive_words")


_EXPLICIT_VALIDATORS = {
    "logger.toml": _validate_logger,
    "output.toml": _validate_output,
    "gui.toml": _validate_gui,
    "conversion.toml": _validate_conversion,
    "export.toml": _validate_export,
    "link.toml": _validate_link,
    "software.toml": _validate_software,
    "optimize.toml": _validate_optimize,
    "field_processors.toml": _validate_field_processors,
    "numbering/add.toml": _validate_numbering_add,
    "numbering/cleanup.toml": _validate_numbering_cleanup,
    "proofread/pairs.toml": _validate_proofread_pairs,
    "proofread/symbol_map.toml": _validate_proofread_symbol_map,
    "proofread/typos.toml": _validate_proofread_typos,
    "proofread/sensitive_words.toml": _validate_proofread_sensitive_words,
}


def validate_config_file(
    rel_path: str,
    effective: object,
    shipped: object,
) -> dict[str, Any]:
    """Return a validated, normalized copy of one effective config file.

    Unknown keys are retained so optimizer IDs, numbering schemes, and future
    extension data remain forwards-compatible.  Values for keys already owned
    by the shipped file must keep the shipped shape and primitive type.
    """
    if not isinstance(effective, Mapping):
        raise ConfigSemanticError(f"{rel_path} root must be a table")
    if not isinstance(shipped, Mapping):
        raise ConfigSemanticError(f"shipped {rel_path} root must be a table")

    normalized = deepcopy(dict(effective))
    _validate_known_shapes(normalized, shipped, path=rel_path)
    validator = _EXPLICIT_VALIDATORS.get(rel_path)
    if validator is not None:
        validator(normalized)
    return normalized
