"""utils 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.md2xlsx.core import process_cell_value
from docwen.utils.link_processing import TABLE_CELL_BR_TOKEN, process_markdown_links

pytestmark = pytest.mark.unit


def test_table_safe_embedded_md_file_flattens_newlines_and_escapes_pipes(tmp_path: Path) -> None:
    embedded = tmp_path / "embed.md"
    embedded.write_text("第一行\n第二行 | 三", encoding="utf-8")

    source = tmp_path / "src.md"
    source_text = "| 列 |\n| --- |\n| ![[embed.md]] |\n"
    source.write_text(source_text, encoding="utf-8")

    result = process_markdown_links(source_text, str(source), table_safe=True)
    assert f"| 第一行{TABLE_CELL_BR_TOKEN}第二行 \\| 三 |" in result

    result_raw = process_markdown_links(source_text, str(source), table_safe=False)
    assert "第一行\n第二行 | 三" in result_raw


def test_md2xlsx_process_cell_value_restores_table_break_token() -> None:
    value, _ = process_cell_value(f"A{TABLE_CELL_BR_TOKEN}B")
    assert value == "A\nB"
