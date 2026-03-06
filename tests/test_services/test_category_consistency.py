"""services 单元测试。"""

from __future__ import annotations

import pytest

from docwen import formats
from docwen.services.strategies import registry
from docwen.utils.file_type_utils import get_strategy_file_category

pytestmark = pytest.mark.unit


def test_strategy_file_category_never_returns_text(tmp_path) -> None:
    p = tmp_path / "sample.txt"
    p.write_text("hello", encoding="utf-8")
    category = get_strategy_file_category(str(p))
    assert category != "text"
    assert category == formats.CATEGORY_MARKDOWN


def test_registry_categories_match_formats_constants() -> None:
    expected = {
        formats.CATEGORY_DOCUMENT,
        formats.CATEGORY_SPREADSHEET,
        formats.CATEGORY_LAYOUT,
        formats.CATEGORY_IMAGE,
        formats.CATEGORY_MARKDOWN,
    }
    assert set(registry._CATEGORY_TO_PACKAGE.keys()) == expected
