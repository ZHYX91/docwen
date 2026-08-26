"""i18n 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_gui import i18n as gui_i18n

pytestmark = pytest.mark.unit


PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOCALES_DIR = PROJECT_ROOT / "i18n" / "locales"


def get_available_locales():
    return [{"code": p.stem} for p in sorted(LOCALES_DIR.glob("*.toml"))]


def t_locale(key: str, locale: str) -> str:
    prev = gui_i18n.get_locale()
    try:
        gui_i18n.set_locale(locale)
        return gui_i18n.t(key, default=f"[{key}]")
    finally:
        gui_i18n.set_locale(prev)


@pytest.mark.parametrize(
    "key",
    [
        "settings.extraction.image_extraction_mode_file",
        "settings.extraction.image_extraction_mode_base64",
        "settings.extraction.ocr_placement_mode_image_md",
        "settings.extraction.ocr_placement_mode_main_md",
        "settings.extraction.ocr_blockquote_title_enabled_label",
        "settings.extraction.ocr_blockquote_title_enabled_tooltip",
    ],
)
def test_settings_extraction_option_keys_exist_in_all_locales(key: str) -> None:
    for locale in get_available_locales():
        code = locale["code"]
        translated = t_locale(key, code)
        assert translated != f"[{key}]"


def test_settings_link_hyperlink_mode_exists_in_all_locales() -> None:
    key = "settings.link.modes.hyperlink"

    for locale in get_available_locales():
        code = locale["code"]
        translated = t_locale(key, code)
        assert translated != f"[{key}]"


def test_settings_link_hyperlink_mode_uses_expected_chinese_wording() -> None:
    assert t_locale("settings.link.modes.hyperlink", "zh_CN") == "超链接"
    assert t_locale("settings.link.modes.hyperlink", "zh_TW") == "超連結"


@pytest.mark.parametrize(
    "key",
    [
        "components.file_drop.batch_list.retry_selected_failed",
        "components.file_drop.batch_list.retry_all_failed",
    ],
)
def test_batch_retry_menu_keys_exist_in_all_locales(key: str) -> None:
    for locale in get_available_locales():
        code = locale["code"]
        translated = t_locale(key, code)
        assert translated != f"[{key}]"
