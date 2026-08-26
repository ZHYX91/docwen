"""Unified i18n: TOML locale loading, nested key lookup & locale-aware helpers."""

from __future__ import annotations

import contextlib
from pathlib import Path
from typing import Any


def locale_is_eligible(locales: object, locale: str) -> bool:
    """Return whether locale metadata admits ``locale``.

    Missing or empty metadata is unrestricted, ``"*"`` is universal, and
    malformed non-sequence metadata fails closed.
    """
    if locales is None:
        return True
    if isinstance(locales, str):
        values = (locales,)
    elif isinstance(locales, (list, tuple, set, frozenset)):
        values = tuple(str(item) for item in locales)
    else:
        return False
    return not values or "*" in values or locale in values


def _resolve_table(data: object, dotted_key: str) -> dict[str, Any]:
    """Resolve a canonical nested TOML table, failing closed."""
    if not isinstance(data, dict):
        return {}

    current: object = data
    for part in dotted_key.split("."):
        if not part or not isinstance(current, dict):
            return {}
        current = current.get(part)
    return current if isinstance(current, dict) else {}


class I18nManager:
    """Unified i18n: loads TOML locale files, provides nested key lookup + locale-aware helpers.

    Use ``set_locale()`` to change the active locale at runtime.
    Call ``clear_cache()`` if locale files are swapped while the application is running.
    """

    def __init__(
        self,
        locales_dir: Path | str,
        default_locale: str = "zh_CN",
    ) -> None:
        self._locales_dir = Path(locales_dir)
        self._default_locale = default_locale
        self._current_locale = default_locale
        self._cache: dict[str, dict] = {}

    # ── locale management ──────────────────────────────────────────

    def set_locale(self, locale: str) -> None:
        """Switch the active locale."""
        self._current_locale = locale

    def get_locale(self) -> str:
        """Return the currently active locale code."""
        return self._current_locale

    def clear_cache(self) -> None:
        """Discard all cached locale file data."""
        self._cache.clear()

    # ── internal loading ───────────────────────────────────────────

    def _load_locale(self, locale: str) -> dict:
        if locale not in self._cache:
            path = self._locales_dir / f"{locale}.toml"
            self._cache[locale] = load_locale_table(path)
        return self._cache[locale]

    # ── generic key lookup ─────────────────────────────────────────

    def t(self, key: str, **kwargs: Any) -> str:
        """Nested key lookup: ``section.sub.key``.

        Supports ``{placeholder}`` style formatting via ``**kwargs``.
        Returns *key* itself when the key or section is not found.
        """
        parts = key.split(".")
        data: Any = self._load_locale(self._current_locale)
        for part in parts:
            if isinstance(data, dict):
                data = data.get(part, key)
            else:
                return key
        result = str(data) if data is not None else key
        if kwargs:
            with contextlib.suppress(KeyError, ValueError):
                result = result.format(**kwargs)
        return result

    # ── locale-aware option filtering (M29) ────────────────────────

    def get_localized_options(self, section_key: str) -> dict[str, Any]:
        """Return options from TOML ``[{section_key}]``, filtered by ``[_locales]`` metadata.

        Keys whose ``_locales.{section_key}.{key}`` list includes neither the
        current locale nor the ``"*"`` wildcard are excluded.  Keys starting
        with ``_`` are skipped.
        """
        data = self._load_locale(self._current_locale)
        section = _resolve_table(data, section_key)
        locales_meta = _resolve_table(data, "_locales")
        section_meta = _resolve_table(locales_meta, section_key)

        result: dict[str, Any] = {}
        for key, value in section.items():
            if key.startswith("_"):
                continue
            key_locales = section_meta.get(key, [])
            if not locale_is_eligible(key_locales, self._current_locale):
                continue
            result[key] = value
        return result

    # ── style format helpers (M30) ─────────────────────────────────

    def get_style_format(self, style_key: str = "body") -> dict[str, Any] | None:
        """Read ``[style_formats.{style_key}]`` for per-locale font/size/indent defaults.

        Returns ``None`` when the key is absent or is not a TOML table.
        """
        data = self._load_locale(self._current_locale)
        style_formats = data.get("style_formats", {})
        if isinstance(style_formats, dict):
            style_format = style_formats.get(style_key)
            return style_format if isinstance(style_format, dict) else None
        return None


def load_locale_table(path: str | Path) -> dict[str, Any]:
    """Read a single locale TOML file as a plain dict.

    File missing or parse failure returns an empty dict (no exception) --
    locale loading is a startup path, and missing files should degrade
    gracefully to the default locale.
    """
    from docwen_runtime.toml_io import read_toml_file

    try:
        return read_toml_file(Path(path))
    except (FileNotFoundError, OSError):
        return {}
    except Exception:
        # Invalid TOML or another parser failure degrades to an empty table.
        return {}


__all__ = ["I18nManager", "load_locale_table", "locale_is_eligible"]
