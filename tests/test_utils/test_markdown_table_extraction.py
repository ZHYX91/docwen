"""utils 单元测试。"""

from __future__ import annotations

import pytest

from docwen.utils.markdown_utils import extract_markdown_tables

pytestmark = pytest.mark.unit


def test_extract_markdown_tables_ignores_fenced_code_blocks() -> None:
    md = "```python\n| A | B |\n|---|---|\n| 1 | 2 |\n```\n\n| H | I |\n|---|---|\n| 3 | 4 |\n"

    tables = extract_markdown_tables(md)

    assert len(tables) == 1
    assert tables[0]["headers"] == ["H", "I"]
    assert tables[0]["rows"] == [["3", "4"]]
