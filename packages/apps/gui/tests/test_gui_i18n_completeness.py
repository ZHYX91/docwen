"""I18n key completeness tests.

Validates that all 11 locale files have the same GUI key set for current GUI
namespaces, following ``docs/configuration.md``.

Policy: every GUI key that exists in zh_CN must exist in all other locales.
Non-GUI keys (CLI, conversion progress messages, style formats) are allowed
to differ since they are not part of the GUI refactor scope.
"""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

LOCALES_DIR = Path(__file__).resolve().parents[4] / "i18n" / "locales"

ALL_LOCALES = [
    "en_US",
    "zh_CN",
    "zh_TW",
    "de_DE",
    "es_ES",
    "fr_FR",
    "ja_JP",
    "ko_KR",
    "pt_BR",
    "ru_RU",
    "vi_VN",
]

GUI_KEY_PREFIXES = (
    "main_window.",
    "components.",
    "conversion_panel.",
    "action_area.",
    "settings.",
    "about.",
    "common.",
    "messages.",
    "editors.",
    "file_types.",
    "file_category.",
    "file_format_validation.",
    "conversion.",
    "components.template_selector",
    "components.template_selector_tabbed",
    "components.file_drop",
    "components.file_selector",
    "components.info_area",
    "components.action_area",
    "components.conversion_panel",
    "info_area.",
)


def _flatten(data: dict, prefix: str = "") -> dict[str, str]:
    result: dict[str, str] = {}
    for key, value in data.items():
        full_key = f"{prefix}.{key}" if prefix else key
        if isinstance(value, dict):
            result.update(_flatten(value, full_key))
        elif value is not None:
            result[full_key] = str(value)
    return result


def _load_locale(locale: str) -> dict[str, str]:
    path = LOCALES_DIR / f"{locale}.toml"
    if not path.exists():
        return {}
    return _flatten(tomllib.loads(path.read_text(encoding="utf-8")))


def _is_gui_key(key: str) -> bool:
    return any(key.startswith(prefix) for prefix in GUI_KEY_PREFIXES)


# ── Tests ───────────────────────────────────────────────────────────────


class TestLocaleKeyCompleteness:
    """All locale files must have identical sets of GUI keys."""

    @pytest.fixture(scope="class")
    def all_tables(self) -> dict[str, dict[str, str]]:
        return {locale: _load_locale(locale) for locale in ALL_LOCALES}

    @pytest.fixture(scope="class")
    def zh_cn_gui_keys(self, all_tables: dict[str, dict[str, str]]) -> set[str]:
        return {k for k in all_tables["zh_CN"] if _is_gui_key(k)}

    @pytest.mark.parametrize(
        "locale",
        [loc for loc in ALL_LOCALES if loc != "zh_CN"],
    )
    def test_locale_has_all_zh_cn_gui_keys(
        self, locale: str, all_tables: dict[str, dict[str, str]], zh_cn_gui_keys: set[str]
    ) -> None:
        locale_keys = set(all_tables[locale].keys())
        missing = zh_cn_gui_keys - locale_keys
        assert not missing, (
            f"{locale} is missing {len(missing)} GUI keys vs zh_CN:\n"
            + "\n".join(f"  - {k}" for k in sorted(missing)[:20])
            + ("\n  ..." if len(missing) > 20 else "")
        )

    def test_all_locales_have_same_gui_key_count(self, all_tables: dict[str, dict[str, str]]) -> None:
        zh_cn_gui = {k for k in all_tables["zh_CN"] if _is_gui_key(k)}
        deviations: list[tuple[str, int, int]] = []
        for locale in ALL_LOCALES:
            if locale == "zh_CN":
                continue
            locale_gui = {k for k in all_tables[locale] if _is_gui_key(k)}
            if len(locale_gui) != len(zh_cn_gui):
                deviations.append((locale, len(locale_gui), len(zh_cn_gui)))
        assert not deviations, "Some locales have different GUI key counts vs zh_CN:\n" + "\n".join(
            f"  {loc}: {cnt} vs {ref}" for loc, cnt, ref in deviations
        )


class TestNoMissingEnUS:
    """en_US is the English fallback — it must have all keys zh_CN has."""

    def test_en_us_not_missing_any_zh_cn_key(self) -> None:
        zh_cn = _load_locale("zh_CN")
        en_us = _load_locale("en_US")
        missing = set(zh_cn.keys()) - set(en_us.keys())
        assert not missing, f"en_US is missing {len(missing)} keys vs zh_CN:\n" + "\n".join(
            f"  - {k}" for k in sorted(missing)
        )


class TestKeyNaming:
    """GUI keys must follow the naming conventions from the i18n plan."""

    def test_settings_tab_keys_exist(self) -> None:
        zh_cn = _load_locale("zh_CN")
        expected = {
            "settings.tabs.general",
            "settings.tabs.text",
            "settings.tabs.proofread",
            "settings.tabs.document",
            "settings.tabs.spreadsheet",
            "settings.tabs.image",
            "settings.tabs.layout",
            "settings.tabs.link",
            "settings.tabs.formatting",
            "settings.tabs.output",
            "settings.tabs.export",
            "settings.tabs.logging",
            "settings.tabs.other",
        }
        missing = expected - set(zh_cn.keys())
        assert not missing, f"Missing settings tab keys: {missing}"

    def test_main_window_keys_exist(self) -> None:
        zh_cn = _load_locale("zh_CN")
        expected = {
            "main_window.window_title",
            "main_window.settings_tooltip",
            "main_window.about_tooltip",
            "main_window.font_size_tooltip",
            "main_window.version_offline",
        }
        missing = expected - set(zh_cn.keys())
        assert not missing, f"Missing main_window keys: {missing}"

    def test_common_keys_exist(self) -> None:
        zh_cn = _load_locale("zh_CN")
        expected = {"common.ok", "common.cancel", "common.apply"}
        missing = expected - set(zh_cn.keys())
        assert not missing, f"Missing common keys: {missing}"

    def test_settings_reset_keys_exist(self) -> None:
        zh_cn = _load_locale("zh_CN")
        expected = {
            "settings.reset.tab_button",
            "settings.reset.all_button",
            "settings.reset.tab_confirm_title",
            "settings.reset.all_confirm_title",
        }
        missing = expected - set(zh_cn.keys())
        assert not missing, f"Missing settings.reset keys: {missing}"

    @pytest.mark.parametrize("locale", ALL_LOCALES)
    def test_general_theme_keys_are_not_nested_under_languages(self, locale: str) -> None:
        table = _load_locale(locale)
        expected = {
            "settings.general.theme_section",
            "settings.general.theme_description",
            "settings.general.theme_label",
            "settings.general.theme_tooltip",
            "settings.general.theme_preview",
            "settings.general.sample_button",
            "settings.general.sample_text",
        }
        misplaced = {f"settings.general.languages.{key.rsplit('.', 1)[-1]}" for key in expected}

        missing = expected - set(table)
        nested = misplaced & set(table)
        assert not missing, f"{locale} missing settings.general theme keys: {sorted(missing)}"
        assert not nested, f"{locale} has theme keys nested under settings.general.languages: {sorted(nested)}"
