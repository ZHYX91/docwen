"""Strict request-owned DOCX style catalog projection."""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable, Mapping
from pathlib import Path
from tomllib import TOMLDecodeError
from typing import Any

from docwen_core.docx_styles import (
    CUSTOM_DOCUMENT_STYLE_KEYS,
    SHIPPED_STYLE_LOCALES,
    DocumentStyleCatalog,
    DocumentStyleFormat,
)
from docwen_runtime.templates.registry import TemplateResolutionError
from docwen_runtime.toml_io import read_toml_file

_LOCALE_CODE = re.compile(r"[A-Za-z]{2,3}_[A-Za-z]{2,4}")
_FORMAT_SOURCE_KEYS: tuple[str, ...] = (
    "body_paragraph",
    "heading_1",
    "heading_2",
    "heading_3_9",
)
_JUSTIFICATIONS = frozenset({"left", "right", "center", "both", "distribute"})
_MISSING = object()


class DocumentStyleCatalogError(TemplateResolutionError):
    """A locale resource cannot satisfy the frozen style contract."""

    def __init__(self, diagnostic_code: str, message: str, *, error_type: str) -> None:
        super().__init__(message, diagnostic_code=diagnostic_code)
        self.error_type = error_type


def build_document_style_catalog(
    config_snapshot: Mapping[str, Any] | None,
    *,
    request_options: Mapping[str, Any] | None = None,
    locales_dir: Path | None = None,
) -> DocumentStyleCatalog:
    """Build one immutable catalog without cross-locale output fallback."""

    snapshot = config_snapshot if isinstance(config_snapshot, Mapping) else {}
    options = request_options if isinstance(request_options, Mapping) else {}
    locale = _resolve_locale(snapshot, options)
    if locale not in SHIPPED_STYLE_LOCALES:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_LOCALE_UNSUPPORTED",
            f"DOCX style locale is not shipped: {locale!r}",
            error_type="invalid_input",
        )

    if locales_dir is None:
        from docwen_runtime.resources import ResourceRegistry

        locales_dir = ResourceRegistry.default().locales_dir()
    locale_root = Path(locales_dir)

    tables = {code: _read_locale_table(locale_root, code) for code in SHIPPED_STYLE_LOCALES}
    output_names = _style_names(tables[locale], locale)
    all_names = {code: _style_names(tables[code], code) for code in SHIPPED_STYLE_LOCALES}
    recognition_names = _recognition_names(all_names)
    formats = _style_formats(tables[locale], locale)
    return DocumentStyleCatalog(
        locale=locale,
        output_names=tuple((key, output_names[key]) for key in CUSTOM_DOCUMENT_STYLE_KEYS),
        recognition_names=recognition_names,
        formats=formats,
    )


def _resolve_locale(snapshot: Mapping[str, Any], options: Mapping[str, Any]) -> str:
    if "locale" in options:
        return _required_locale(options["locale"], source="request option")

    gui = _mapping(snapshot.get("gui"))
    language = gui.get("language", _MISSING)
    if isinstance(language, Mapping) and "locale" in language:
        return _required_locale(language["locale"], source="configuration")
    if language is not _MISSING and not isinstance(language, Mapping):
        return _required_locale(language, source="configuration")
    return "zh_CN"


def _required_locale(value: object, *, source: str) -> str:
    locale = value.strip() if isinstance(value, str) else ""
    if not locale or _LOCALE_CODE.fullmatch(locale) is None:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_LOCALE_INVALID",
            f"DOCX style locale from {source} must be a non-empty locale code.",
            error_type="invalid_input",
        )
    return locale


def _read_locale_table(locales_dir: Path, locale: str) -> Mapping[str, Any]:
    path = locales_dir / f"{locale}.toml"
    try:
        return read_toml_file(path)
    except FileNotFoundError as exc:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_LOCALE_MISSING",
            f"DOCX style locale resource is missing: {locale}",
            error_type="conversion_failed",
        ) from exc
    except TOMLDecodeError as exc:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_LOCALE_INVALID_TOML",
            f"DOCX style locale resource is invalid TOML: {locale}",
            error_type="conversion_failed",
        ) from exc
    except (OSError, UnicodeError) as exc:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_LOCALE_UNREADABLE",
            f"DOCX style locale resource is unreadable: {locale}",
            error_type="conversion_failed",
        ) from exc
    except Exception as exc:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_LOCALE_INVALID_TOML",
            f"DOCX style locale resource is invalid TOML: {locale}",
            error_type="conversion_failed",
        ) from exc


def _style_names(table: Mapping[str, Any], locale: str) -> dict[str, str]:
    raw_styles = table.get("styles")
    if not isinstance(raw_styles, Mapping):
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_NAMES_MISSING",
            f"DOCX style locale has no [styles] table: {locale}",
            error_type="conversion_failed",
        )
    actual_keys = tuple(str(key) for key in raw_styles)
    if set(actual_keys) != set(CUSTOM_DOCUMENT_STYLE_KEYS):
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_NAMES_INCOMPLETE",
            f"DOCX style locale keys do not match the 27-style contract: {locale}",
            error_type="conversion_failed",
        )

    names: dict[str, str] = {}
    normalized: dict[str, str] = {}
    for key in CUSTOM_DOCUMENT_STYLE_KEYS:
        raw_name = raw_styles.get(key)
        name = raw_name.strip() if isinstance(raw_name, str) else ""
        if not name:
            raise DocumentStyleCatalogError(
                "DOCX_STYLE_NAME_BLANK",
                f"DOCX style name {key!r} is blank for locale {locale}.",
                error_type="conversion_failed",
            )
        identity = _normalized_name(name)
        if identity in normalized:
            raise DocumentStyleCatalogError(
                "DOCX_STYLE_NAME_DUPLICATE",
                f"DOCX style names {normalized[identity]!r} and {key!r} collide for locale {locale}.",
                error_type="conversion_failed",
            )
        names[key] = name
        normalized[identity] = key
    return names


def _recognition_names(
    all_names: Mapping[str, Mapping[str, str]],
) -> tuple[tuple[str, tuple[str, ...]], ...]:
    ownership: dict[str, str] = {}
    result: list[tuple[str, tuple[str, ...]]] = []
    for key in CUSTOM_DOCUMENT_STYLE_KEYS:
        values = _unique_names(all_names[code][key] for code in SHIPPED_STYLE_LOCALES)
        for value in values:
            identity = _normalized_name(value)
            owner = ownership.setdefault(identity, key)
            if owner != key:
                raise DocumentStyleCatalogError(
                    "DOCX_STYLE_RECOGNITION_AMBIGUOUS",
                    f"DOCX style recognition name {value!r} belongs to both {owner!r} and {key!r}.",
                    error_type="conversion_failed",
                )
        result.append((key, values))
    return tuple(result)


def _unique_names(values: Iterable[str]) -> tuple[str, ...]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        identity = _normalized_name(value)
        if identity not in seen:
            result.append(value)
            seen.add(identity)
    return tuple(result)


def _style_formats(table: Mapping[str, Any], locale: str) -> tuple[tuple[str, DocumentStyleFormat], ...]:
    raw_formats = table.get("style_formats")
    if not isinstance(raw_formats, Mapping) or {str(key) for key in raw_formats} != set(_FORMAT_SOURCE_KEYS):
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_FORMATS_INCOMPLETE",
            f"DOCX style formats do not match the required contract: {locale}",
            error_type="conversion_failed",
        )

    source = {key: _style_format(raw_formats.get(key), locale=locale, key=key) for key in _FORMAT_SOURCE_KEYS}
    return (
        ("body_paragraph", source["body_paragraph"]),
        ("heading_1", source["heading_1"]),
        ("heading_2", source["heading_2"]),
        *((f"heading_{level}", source["heading_3_9"]) for level in range(3, 10)),
    )


def _style_format(value: object, *, locale: str, key: str) -> DocumentStyleFormat:
    if not isinstance(value, Mapping):
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_FORMAT_INVALID",
            f"DOCX style format {key!r} is not a table for locale {locale}.",
            error_type="conversion_failed",
        )
    expected = {
        "east_asia_font",
        "ascii_font",
        "font_size_pt",
        "first_line_indent_chars",
        "first_line_indent_cm",
        "spacing_after_twip",
        "spacing_before_twip",
        "bold",
        "justification",
    }
    if set(value) != expected:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_FORMAT_INVALID",
            f"DOCX style format {key!r} has an invalid field set for locale {locale}.",
            error_type="conversion_failed",
        )
    east_asia_font = _nonblank_string(value.get("east_asia_font"))
    ascii_font = _nonblank_string(value.get("ascii_font"))
    justification = _nonblank_string(value.get("justification"))
    if not east_asia_font or not ascii_font or justification not in _JUSTIFICATIONS:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_FORMAT_INVALID",
            f"DOCX style format {key!r} contains invalid text values for locale {locale}.",
            error_type="conversion_failed",
        )
    font_size_pt = _positive_number(value.get("font_size_pt"))
    first_line_chars = _integer(value.get("first_line_indent_chars"))
    first_line_cm = _number(value.get("first_line_indent_cm"))
    spacing_after = _integer(value.get("spacing_after_twip"))
    spacing_before = _integer(value.get("spacing_before_twip"))
    bold = value.get("bold")
    if (
        font_size_pt is None
        or first_line_chars is None
        or first_line_cm is None
        or spacing_after is None
        or spacing_before is None
        or not isinstance(bold, bool)
    ):
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_FORMAT_INVALID",
            f"DOCX style format {key!r} contains invalid values for locale {locale}.",
            error_type="conversion_failed",
        )
    return DocumentStyleFormat(
        east_asia_font=east_asia_font,
        ascii_font=ascii_font,
        font_size_pt=font_size_pt,
        first_line_indent_chars=first_line_chars,
        first_line_indent_cm=first_line_cm,
        spacing_after_twip=spacing_after,
        spacing_before_twip=spacing_before,
        bold=bold,
        justification=justification,
    )


def _mapping(value: object) -> Mapping[str, Any]:
    return value if isinstance(value, Mapping) else {}


def _nonblank_string(value: object) -> str:
    return value.strip() if isinstance(value, str) else ""


def _number(value: object) -> float | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return float(value)


def _positive_number(value: object) -> float | None:
    number = _number(value)
    return number if number is not None and number > 0 else None


def _integer(value: object) -> int | None:
    return value if isinstance(value, int) and not isinstance(value, bool) else None


def _normalized_name(value: str) -> str:
    return unicodedata.normalize("NFC", value).casefold()


__all__ = ["DocumentStyleCatalogError", "build_document_style_catalog"]
