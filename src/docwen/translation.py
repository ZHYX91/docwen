from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

_translator: Callable[[str, str | None], str] | None = None
_locale_tables: dict[str, dict[str, Any]] | None = None


def set_translator(translator: Callable[[str, str | None], str] | None) -> None:
    global _translator
    _translator = translator


def t(key: str, default: str | None = None, **kwargs: Any) -> str:
    if _translator is not None:
        try:
            return _translator(key, default, **kwargs)
        except TypeError:
            return _translator(key, default)

    text = default if default is not None else key
    try:
        return str(text).format(**kwargs)
    except Exception:
        return str(text)


def _load_locale_tables() -> dict[str, dict[str, Any]]:
    global _locale_tables
    if _locale_tables is not None:
        return _locale_tables
    try:
        import tomllib
    except Exception:  # pragma: no cover
        import tomli as tomllib  # type: ignore[no-redef]

    base = Path(__file__).resolve().parent / "i18n" / "locales"
    tables: dict[str, dict[str, Any]] = {}
    for path in base.glob("*.toml"):
        try:
            tables[path.stem] = tomllib.loads(path.read_text(encoding="utf-8"))
        except Exception:
            tables[path.stem] = {}
    _locale_tables = tables
    return tables


def t_all_locales(key: str) -> dict[str, str]:
    parts = [p for p in str(key).split(".") if p]
    if not parts:
        return {}

    tables = _load_locale_tables()
    out: dict[str, str] = {}
    for locale, table in tables.items():
        value: Any = table
        for part in parts:
            if not isinstance(value, dict) or part not in value:
                value = None
                break
            value = value.get(part)
        if isinstance(value, str):
            out[locale] = value
    return out


def t_locale(key: str, locale: str, default: str | None = None, **kwargs: Any) -> str:
    parts = [p for p in str(key).split(".") if p]
    if not parts:
        return t(key, default=default, **kwargs)

    tables = _load_locale_tables()
    table = tables.get(str(locale)) or {}
    value: Any = table
    for part in parts:
        if not isinstance(value, dict) or part not in value:
            value = None
            break
        value = value.get(part)
    text = value if isinstance(value, str) else (default if default is not None else key)
    try:
        return str(text).format(**kwargs)
    except Exception:
        return str(text)


def get_current_locale() -> str:
    try:
        from docwen.i18n import get_current_locale as _get_current_locale

        return str(_get_current_locale() or "zh_CN")
    except Exception:
        return "zh_CN"


def get_style_format(style_key: str, locale: str | None = None) -> dict[str, Any]:
    tables = _load_locale_tables()
    loc = str(locale or get_current_locale() or "zh_CN")
    table = tables.get(loc) or {}
    style_formats = table.get("style_formats")
    if not isinstance(style_formats, dict):
        return {}
    fmt = style_formats.get(style_key)
    return fmt if isinstance(fmt, dict) else {}
