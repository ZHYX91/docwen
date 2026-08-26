from __future__ import annotations

import threading
from pathlib import Path
from typing import Any

from docwen_runtime.i18n import load_locale_table

from .resources import i18n_locales_dir

_DEFAULT_LOCALE = "zh_CN"
_current_locale = _DEFAULT_LOCALE
_lock = threading.Lock()
_cache: dict[str, dict[str, str]] = {}


def set_locale(locale: str) -> None:
    global _current_locale
    if locale:
        _current_locale = locale


def get_locale() -> str:
    return _current_locale


def _flatten(prefix: str, value: Any, out: dict[str, str]) -> None:
    if isinstance(value, dict):
        for k, v in value.items():
            next_prefix = f"{prefix}.{k}" if prefix else str(k)
            _flatten(next_prefix, v, out)
        return
    if value is None:
        return
    out[prefix] = str(value)


def _load_locale(locale: str) -> dict[str, str]:
    with _lock:
        cached = _cache.get(locale)
        if cached is not None:
            return cached

        locale_dir = i18n_locales_dir()
        file_path = Path(locale_dir) / f"{locale}.toml"

        if not file_path.exists() and locale != _DEFAULT_LOCALE:
            file_path = Path(locale_dir) / f"{_DEFAULT_LOCALE}.toml"

        data: dict[str, Any] = {}
        if file_path.exists():
            data = load_locale_table(file_path)

        flat: dict[str, str] = {}
        _flatten("", data, flat)
        _cache[locale] = flat
        return flat


def t(key: str, default: str = "", **kwargs: object) -> str:
    locale = get_locale()
    table = _load_locale(locale)
    value = table.get(key)
    text = value if value is not None else default
    if kwargs:
        try:
            return str(text).format(**kwargs)
        except Exception:
            return str(text)
    return str(text)
