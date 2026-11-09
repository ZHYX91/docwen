"""Native quote paragraphs retain their levels across repeated conversions."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from markdown_it import MarkdownIt

from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from tests.integration._round_trip_helper import docx_to_md, md_to_docx

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def _quote_paragraphs(markdown: str) -> list[tuple[int, str]]:
    if markdown.startswith("---\n"):
        markdown = markdown.split("\n---\n", 1)[1]
    paragraphs: list[tuple[int, str]] = []

    def inline_text(node: dict[str, Any]) -> str:
        kind = node["type"]
        if kind in {"inline_math", "block_math"}:
            return f"[{kind}]"
        if kind in {"softbreak", "linebreak"} or (kind == "inline_html" and node.get("raw") == "<br>"):
            return "\n"
        if "children" in node:
            return "".join(inline_text(child) for child in node["children"])
        return node.get("raw", "")

    def visit(nodes: list[dict[str, Any]], depth: int = 0) -> None:
        for node in nodes:
            kind = node["type"]
            if kind == "block_quote":
                visit(node["children"], depth + 1)
            elif depth and kind == "paragraph":
                paragraphs.append((depth, inline_text(node)))
            elif "children" in node:
                visit(node["children"], depth)

    visit(parse_markdown_text(markdown))
    return paragraphs


@pytest.mark.parametrize(
    ("source", "levels"),
    [
        (
            "> 这是一级引用，包含**粗体**和*斜体*格式。\n\n"
            ">> 这是二级引用，包含`代码`和~~删除线~~。\n\n"
            ">>> 引用会自动应用文档模板中的“引用 1~9”段落样式。如果模板中没有此段落样式，程序会自动注入。\n\n"
            "> 引用中的公式：$E=mc^2$，以及**粗体公式描述**。\n",
            [1, 2, 3, 1],
        ),
        ("> First paragraph.\n>\n> Second paragraph.\n", [1, 1]),
        ("> Parent.\n>\n>> Child.\n>>\n>>> Deep.\n>\n> Parent again.\n", [1, 2, 3, 1]),
        ("> First line.<br>Second line.\n", [1]),
    ],
    ids=["original-four-paragraphs", "same-level-paragraphs", "nested-level-return", "one-native-paragraph-with-break"],
)
def test_quote_paragraph_levels_and_boundaries_survive_two_round_trips(
    native_round_trip_runtime: Any, tmp_path: Path, source: str, levels: list[int]
) -> None:
    expected = _quote_paragraphs(source)
    assert [depth for depth, _ in expected] == levels
    source_path = tmp_path / "quotes.md"
    source_path.write_text(source, encoding="utf-8")
    current = md_to_docx(native_round_trip_runtime, source_path, tmp_path / "first")
    source_path.unlink()

    for generation in range(2):
        returned = docx_to_md(native_round_trip_runtime, current, tmp_path / f"reverse-{generation}")
        actual = _quote_paragraphs(returned)
        assert [depth for depth, _ in actual] == levels
        assert actual == expected
        if generation == 0:
            next_source = tmp_path / "second.md"
            next_source.write_text(returned, encoding="utf-8")
            current = md_to_docx(native_round_trip_runtime, next_source, tmp_path / "second")
            next_source.unlink()


def test_quote_list_stays_one_container_and_one_list_across_two_round_trips(
    native_round_trip_runtime: Any, tmp_path: Path
) -> None:
    source = "> 1. First item.\n> 2. Second item.\n\n^quoted-list\n"
    source_path = tmp_path / "quote-list.md"
    source_path.write_text(source, encoding="utf-8")
    current = md_to_docx(native_round_trip_runtime, source_path, tmp_path / "first-list")
    source_path.unlink()

    for generation in range(2):
        returned = docx_to_md(native_round_trip_runtime, current, tmp_path / f"list-reverse-{generation}")
        body = returned.split("\n---\n", 1)[1] if returned.startswith("---\n") else returned
        tokens = MarkdownIt("commonmark").parse(body)
        assert [
            (token.type, token.level) for token in tokens if token.type in {"blockquote_open", "ordered_list_open"}
        ] == [
            ("blockquote_open", 0),
            ("ordered_list_open", 1),
        ]
        assert len([token for token in tokens if token.type == "list_item_open"]) == 2
        assert [token.content for token in tokens if token.type == "inline" and token.content != "^quoted-list"] == [
            "First item.",
            "Second item.",
        ]
        assert returned.count("^quoted-list") == 1
        if generation == 0:
            next_source = tmp_path / "quote-list-second.md"
            next_source.write_text(returned, encoding="utf-8")
            current = md_to_docx(native_round_trip_runtime, next_source, tmp_path / "second-list")
            next_source.unlink()
