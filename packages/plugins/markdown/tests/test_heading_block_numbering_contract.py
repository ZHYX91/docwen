"""Source heading edits must preserve the production block grammar."""

from __future__ import annotations

import pytest

from docwen_core.links import markdown_heading_ranges, markdown_source_context_lines
from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_core.text.heading_numbering import HeadingFormatter
from docwen_plugin_markdown.common_utils import _apply_numbering
from docwen_plugin_markdown.mistune_extensions import create_extended_markdown

pytestmark = pytest.mark.contract


def _formatter() -> HeadingFormatter:
    return HeadingFormatter({f"level_{level}": {"format": f"{{{level}.arabic_half}} "} for level in range(1, 10)})


@pytest.mark.parametrize("marker", ["0.", "2.", "9)", "123."])
def test_numbering_keeps_noninterrupting_list_spellings_in_their_paragraph(marker: str) -> None:
    source = f"intro\n{marker} # Fake\n\n# Real\nBody\n"
    numbered = _apply_numbering(source, _formatter())
    assert numbered == f"intro\n{marker} # Fake\n\n# 1 Real\nBody\n"
    ast = create_extended_markdown(extensions=MarkdownExtensions())(numbered)
    assert isinstance(ast, list)
    assert [node["type"] for node in ast] == ["paragraph", "blank_line", "heading", "paragraph"]


@pytest.mark.parametrize("rule", ["***", "* * *", "___", "_ _ _", "---"])
def test_numbering_preserves_a_rule_followed_by_plain_equals(rule: str) -> None:
    source = f"{rule}\n===\n\n# Real\nBody\n"
    numbered = _apply_numbering(source, _formatter())
    assert numbered == f"{rule}\n===\n\n# 1 Real\nBody\n"
    ast = create_extended_markdown(extensions=MarkdownExtensions())(numbered)
    assert isinstance(ast, list)
    assert [node["type"] for node in ast] == ["thematic_break", "paragraph", "blank_line", "heading", "paragraph"]


def test_numbering_edits_the_first_line_of_a_complete_multiline_setext_title() -> None:
    source = "foo\n2. bar\n===\n\n# Real\n"
    assert _apply_numbering(source, _formatter()) == "1 foo\n2. bar\n===\n\n# 2 Real\n"


@pytest.mark.parametrize("level", range(1, 10))
@pytest.mark.parametrize("prefix", ["", "> ", "- "])
@pytest.mark.parametrize("closing", ["", " ###"])
def test_numbering_preserves_empty_atx_separator_and_closing_hashes(level: int, prefix: str, closing: str) -> None:
    source = "\ufeff" + prefix + "#" * level + closing + "\r\n"
    numbered = _apply_numbering(source, _formatter(), extended_headings=True)
    expected = "\ufeff" + prefix + "#" * level + " 1 " + ("###" if closing else "") + "\r\n"
    assert numbered == expected


def _parsed_headings(nodes: list[dict]) -> list[tuple[str, int, str]]:
    def inline_text(children: list[dict]) -> str:
        return "".join(
            "\n"
            if child["type"] in {"softbreak", "linebreak"}
            else child.get("raw", inline_text(child.get("children", [])))
            for child in children
        )

    found: list[tuple[str, int, str]] = []
    for node in nodes:
        if node["type"] == "heading":
            found.append((node["style"], node["attrs"]["level"], inline_text(node.get("children", []))))
        else:
            found.extend(_parsed_headings(node.get("children", [])))
    return found


@pytest.mark.parametrize(
    "source, expected",
    [
        ("intro\n2. # Fake\n\n# Real\n", [("atx", 1, "Real")]),
        ("intro\n1. # One\n2. # Two\n3. # Three\n", [("atx", 1, "One"), ("atx", 1, "Two"), ("atx", 1, "Three")]),
        ("foo\n2. bar\n===\n", [("setext", 1, "foo\n2. bar")]),
        ("***\n===\n\n# Real\n", [("atx", 1, "Real")]),
        ("- text\nlazy\n  ```text\n  # Hidden\n\n# Real\n", [("atx", 1, "Real")]),
        ("- parent\n  - text\nlazy\n    ```text\n    # Hidden\n\n# Real\n", [("atx", 1, "Real")]),
        ("> text\nlazy\n> ```text\n> # Hidden\n\n# Real\n", [("atx", 1, "Real")]),
        ("- a\n===\n", [("setext", 1, "a")]),
        ("> a\n===\n", []),
        ("> # H\ntext\n", [("atx", 1, "H")]),
        ("> intro\n# H\ntext\n", [("atx", 1, "H")]),
    ],
)
def test_core_heading_inventory_and_real_parser_share_block_qualification(
    source: str, expected: list[tuple[str, int, str]]
) -> None:
    parser = create_extended_markdown(extensions=MarkdownExtensions())
    ast = parser(source)
    assert isinstance(ast, list)
    assert _parsed_headings(ast) == expected
    assert [
        (heading.style, heading.level, "\n".join(source[start:end] for start, end in heading.text_ranges))
        for heading in markdown_heading_ranges(source, extended_headings=False)
    ] == expected


@pytest.mark.parametrize("depth", [19, 20, 21])
@pytest.mark.parametrize("prefix", ["> ", "- "])
def test_heading_inventory_stops_at_the_production_container_depth(depth: int, prefix: str) -> None:
    source = prefix * depth + "# Title\n"
    expected = [("atx", 1, "Title")] if depth <= 20 else []
    parser = create_extended_markdown(extensions=MarkdownExtensions())
    ast = parser(source)
    assert isinstance(ast, list)
    assert _parsed_headings(ast) == expected
    assert [(h.style, h.level, source[h.text_start : h.text_end]) for h in markdown_heading_ranges(source)] == expected


@pytest.mark.parametrize("prefix", ["", "> ", "- "])
@pytest.mark.parametrize("first", ["#", "# ###", "# "])
def test_empty_atx_counts_once_and_disabled_formatting_preserves_bytes(prefix: str, first: str) -> None:
    source = prefix + first + "\r\n\r\n# Tail\r\n"
    numbered = _apply_numbering(source, _formatter())
    parser = create_extended_markdown(extensions=MarkdownExtensions())
    ast = parser(numbered)
    assert isinstance(ast, list)
    assert _parsed_headings(ast) == [("atx", 1, "1"), ("atx", 1, "2 Tail")]
    disabled = HeadingFormatter({f"level_{level}": {"format": ""} for level in range(1, 10)})
    assert _apply_numbering(source, disabled) == source


@pytest.mark.parametrize(
    "source, expected_math",
    [
        ("> $$x$$\n", ["x"]),
        ("- $$x$$\n", ["x"]),
        ("> - $$\n>   # Literal ^fake\n>   $$\n", ["# Literal ^fake"]),
        ("- > $$\n  > # Literal ^fake\n  > $$\n", ["# Literal ^fake"]),
        ("> $$\n# Literal ^fake\n$$\n", ["# Literal ^fake"]),
        ("- $$\n# Real\n$$\n", []),
        ("- $$\nbody\n$$\n", ["body"]),
        ("- $$\n  body\n\n# Real\n$$\n", []),
        ("- $$\n  body\n  - $$\n", []),
        ("$$\n$$\n", []),
        ("$$\n\n$$\n", [""]),
        ("$$$$\n", [""]),
        ("$$\nbody\n $$\n", []),
    ],
)
def test_display_math_inventory_matches_actual_block_parser_container_scope(
    source: str, expected_math: list[str]
) -> None:
    def math_blocks(nodes: list[dict]) -> list[str]:
        found: list[str] = []
        for node in nodes:
            if node["type"] == "block_math":
                found.append(node["raw"])
            elif node["type"] in {"block_quote", "list", "list_item"}:
                found.extend(math_blocks(node["children"]))
        return found

    parser = create_extended_markdown(extensions=MarkdownExtensions())
    ast = parser(source)
    assert isinstance(ast, list)
    assert math_blocks(ast) == expected_math
    rows = markdown_source_context_lines(source)
    assert sum(row.block_kind == "math_block" for row in rows) == len(expected_math)
    if expected_math:
        assert markdown_heading_ranges(source) == []
