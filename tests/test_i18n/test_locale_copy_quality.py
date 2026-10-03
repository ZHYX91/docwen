"""Regression guards for reviewed native-language UI labels."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

_LOCALES = Path(__file__).resolve().parents[2] / "i18n" / "locales"


def _value(locale: str, section: str, key: str) -> str:
    data = tomllib.loads((_LOCALES / f"{locale}.toml").read_text(encoding="utf-8"))
    node = data
    for part in section.split("."):
        node = node[part]
    value = node[key]
    assert isinstance(value, str)
    return value


@pytest.mark.parametrize(
    ("locale", "section", "key", "expected"),
    [
        ("ru_RU", "components.file_drop", "add_file_action", "Добавить файлы"),
        ("ru_RU", "components.file_drop", "add_folder_action", "Добавить папку"),
        ("ru_RU", "components.file_drop", "select_folder_dialog", "Выбрать папку"),
        ("zh_TW", "components.template_selector", "source_tooltip", "來源資料夾：{value}"),
        ("vi_VN", "components.file_drop", "add_file_action", "Thêm tệp"),
        ("vi_VN", "components.file_drop", "add_folder_action", "Thêm thư mục"),
        ("vi_VN", "components.file_drop.batch_list", "filter_button_with_label", "Lọc: {label}"),
        ("vi_VN", "main_window", "settings_tooltip", "Cài đặt (Ctrl+,)"),
    ],
)
def test_reviewed_visible_labels_remain_native(
    locale: str,
    section: str,
    key: str,
    expected: str,
) -> None:
    assert _value(locale, section, key) == expected
