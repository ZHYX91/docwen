"""Independent Markdown table consumer for native spreadsheet value assertions."""

from __future__ import annotations

import csv
import html
import io
from typing import Any

import mistune


def _text(nodes: list[dict[str, Any]]) -> str:
    pieces: list[str] = []
    for node in nodes:
        if node["type"] in {"text", "codespan"}:
            pieces.append(html.unescape(node.get("raw", "")))
        elif node["type"] == "inline_html" and node.get("raw", "").lower() in {"<br>", "<br/>", "<br />"}:
            pieces.append("\n")
        elif "children" in node:
            pieces.append(_text(node["children"]))
    return "".join(pieces)


def markdown_table_values(markdown: str) -> list[list[list[str]]]:
    """Read GFM tables without any DocWen encoder, parser or serializer."""
    ast = mistune.create_markdown(renderer="ast", plugins=["table"])(markdown)
    assert isinstance(ast, list)
    matrices: list[list[list[str]]] = []
    for node in ast:
        if node["type"] != "table":
            continue
        rows = []
        for section in node["children"]:
            if section["type"] == "table_head":
                rows.append([_text(cell["children"]) for cell in section["children"]])
            else:
                rows.extend([_text(cell["children"]) for cell in row["children"]] for row in section["children"])
        matrices.append(rows)
    return matrices


def assert_old_workbook_source_values(markdown: str, fixture: dict[str, Any]) -> None:
    """Expected matrices come from the recorded input, not its old spacing."""
    expected = []
    for sheet in fixture["input_xlsx"]["sheets"]:
        table = sheet["table"]
        expected.append([table["headers"], *table["rows"]])
        if "merged_block" in sheet:
            expected.append([sheet["merged_block"]])
    assert markdown_table_values(markdown) == expected


def delimited_source_blocks(text: str, delimiter: str) -> list[list[list[str]]]:
    """Independent input facts for fixtures separated by physical empty rows."""
    blocks: list[list[list[str]]] = []
    current: list[list[str]] = []
    for row in csv.reader(io.StringIO(text), delimiter=delimiter):
        if not any(row):
            if current:
                blocks.append(current)
                current = []
        else:
            current.append(row)
    if current:
        blocks.append(current)
    return blocks
