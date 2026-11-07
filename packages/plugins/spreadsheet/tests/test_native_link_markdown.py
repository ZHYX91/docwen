"""Native external links survive reload and Markdown export without active schemes."""

from pathlib import Path

import openpyxl
import pytest

from docwen_plugin_spreadsheet.native_links import native_cell_markdown
from docwen_plugin_spreadsheet.to_markdown.converter import SpreadsheetToMarkdownConverter

from ._xlsx_to_md_golden_support import _build_fake_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    "target",
    ["https://example.com/a b?q=x|y", "mailto:a@example.com", "file:///tmp/example.md", "obsidian://open?vault=Demo"],
)
def test_native_link_keeps_label_and_target_after_reload(tmp_path: Path, target: str) -> None:
    source = tmp_path / "links.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "Link"
    sheet["A2"] = "[Label]|literal"
    sheet["A2"].hyperlink = target
    workbook.save(source)
    original = source.read_bytes()
    stage = tmp_path / "stage"
    stage.mkdir()
    result = SpreadsheetToMarkdownConverter().convert(
        _build_fake_context(str(source), str(stage), source_format="xlsx")
    )
    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    from urllib.parse import quote

    assert f"[\\[Label\\]\\|literal](<{quote(target, safe=':/#?=&%+@;,~-._!')}>)" in markdown
    assert source.read_bytes() == original


@pytest.mark.parametrize(
    "target",
    ["javascript:alert(1)", "data:text/plain,hi", "//example.com/path", "relative.md", "#Sheet1!A1", "https://"],
)
def test_unproven_targets_remain_visible_text(target: str) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "Safe label"
    sheet["A1"].hyperlink = target
    assert native_cell_markdown(sheet["A1"]) == "Safe label"
