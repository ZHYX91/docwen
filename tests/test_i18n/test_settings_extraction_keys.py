"""i18n 单元测试。"""

from __future__ import annotations

import pytest

from docwen.i18n import get_available_locales, t_locale

pytestmark = pytest.mark.unit


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
