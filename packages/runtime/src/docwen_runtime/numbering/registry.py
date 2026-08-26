"""Numbering scheme registry backed by an explicit immutable config snapshot."""

from __future__ import annotations

from collections.abc import Mapping
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docwen_runtime.i18n import load_locale_table, locale_is_eligible


@dataclass(frozen=True)
class NumberingSchemeInfo:
    scheme_id: str
    name: str
    description: str
    enabled: bool
    is_system: bool
    locales: tuple[str, ...]
    levels: dict[str, str]

    def to_dict(self) -> dict[str, object]:
        return {
            "id": self.scheme_id,
            "name": self.name,
            "description": self.description,
            "enabled": self.enabled,
            "is_system": self.is_system,
            "locales": list(self.locales),
            "levels": dict(self.levels),
        }


class NumberingSchemeNotFoundError(LookupError):
    pass


class NumberingSchemeRegistry:
    """List and resolve heading numbering schemes from canonical config."""

    def __init__(
        self,
        *,
        config_snapshot: Mapping[str, Any],
        locale_path: Path | str,
        locale: str | None = None,
    ) -> None:
        self._config_snapshot = deepcopy(dict(config_snapshot))
        self._locale_path = Path(locale_path)
        self._locale = str(locale or _snapshot_locale(config_snapshot) or self._locale_path.stem)

    @classmethod
    def from_config_snapshot(
        cls,
        config_snapshot: Mapping[str, Any],
        *,
        locale_path: Path | str,
    ) -> NumberingSchemeRegistry:
        """Build a registry from dependencies selected by the composition root."""

        return cls(
            config_snapshot=config_snapshot,
            locale_path=locale_path,
            locale=Path(locale_path).stem,
        )

    def with_config_snapshot(
        self,
        config_snapshot: Mapping[str, Any],
        *,
        locale: str | None = None,
    ) -> NumberingSchemeRegistry:
        """Return an isolated request registry while retaining locale ownership."""

        requested_locale = locale or _snapshot_locale(config_snapshot)
        locale_path = self._locale_path
        if requested_locale:
            candidate = self._locale_path.parent / f"{requested_locale}.toml"
            if candidate.is_file():
                locale_path = candidate
        return type(self)(
            config_snapshot=config_snapshot,
            locale_path=locale_path,
            locale=requested_locale,
        )

    def list_schemes(self) -> list[NumberingSchemeInfo]:
        data = self._load_data()
        settings = data.get("settings", {})
        order = list(settings.get("order", []))
        schemes: dict[str, Any] = data.get("schemes", {})

        scheme_ids = order + [scheme_id for scheme_id in schemes if scheme_id not in order]
        translations = self._load_translations()
        result: list[NumberingSchemeInfo] = []
        for scheme_id in scheme_ids:
            raw = schemes.get(scheme_id)
            if not isinstance(raw, dict):
                continue
            if not locale_is_eligible(raw.get("locales", ["*"]), self._locale):
                continue
            result.append(_parse_scheme(scheme_id, raw, translations))
        return result

    def get_scheme(self, scheme_id: str) -> NumberingSchemeInfo:
        requested = scheme_id.strip()
        for scheme in self.list_schemes():
            if scheme.scheme_id == requested:
                return scheme
        raise NumberingSchemeNotFoundError(f"编号方案不存在: {scheme_id}")

    def _load_data(self) -> dict[str, Any]:
        numbering = self._config_snapshot.get("numbering", {})
        data = numbering.get("add", {}) if isinstance(numbering, Mapping) else {}
        return dict(data) if isinstance(data, Mapping) else {}

    def _load_translations(self) -> dict[str, str]:
        if not self._locale_path.is_file():
            return {}
        locale = load_locale_table(self._locale_path)
        translations: dict[str, str] = {}
        for section_path in (
            ("cli", "numbering_schemes"),
            ("editors", "numbering_add", "names"),
            ("editors", "numbering_add", "descriptions"),
        ):
            section = _nested_get(locale, section_path)
            if isinstance(section, dict):
                for key, value in section.items():
                    translations[str(key)] = str(value)
        return translations


def _snapshot_locale(config_snapshot: Mapping[str, Any]) -> str | None:
    gui = config_snapshot.get("gui", {})
    language = gui.get("language", {}) if isinstance(gui, Mapping) else {}
    locale = language.get("locale") if isinstance(language, Mapping) else None
    return str(locale) if locale else None


def _nested_get(data: dict[str, Any], path: tuple[str, ...]) -> Any:
    current: Any = data
    for key in path:
        if not isinstance(current, dict):
            return None
        current = current.get(key)
    return current


def _parse_scheme(scheme_id: str, raw: dict[str, Any], translations: dict[str, str]) -> NumberingSchemeInfo:
    levels: dict[str, str] = {}
    for key, value in raw.items():
        if key.startswith("level_") and isinstance(value, dict):
            levels[key] = str(value.get("format", ""))

    return NumberingSchemeInfo(
        scheme_id=scheme_id,
        name=_display_text(raw.get("name"), raw.get("name_key"), scheme_id, translations),
        description=_display_text(raw.get("description"), raw.get("description_key"), "", translations),
        enabled=bool(raw.get("enabled", True)),
        is_system=bool(raw.get("is_system", False)),
        locales=_normalized_locales(raw.get("locales", ["*"])),
        levels=levels,
    )


def _display_text(value: Any, key: Any, fallback: str, translations: dict[str, str]) -> str:
    if value:
        return str(value)
    if key:
        return translations.get(str(key), str(key))
    return fallback


def _normalized_locales(value: object) -> tuple[str, ...]:
    if isinstance(value, str):
        return (value,)
    if isinstance(value, (list, tuple, set, frozenset)):
        return tuple(str(item) for item in value)
    return ()
