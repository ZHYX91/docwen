"""Setext preprocessing is structural and never rewrites literal code."""

from __future__ import annotations

import pytest

from docwen_plugin_markdown.preprocessor import handle_setext_headings, normalize_html_tags

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("newline", ["\n", "\r\n"])
@pytest.mark.parametrize("trailing", ["", "   ", "\t"])
def test_blank_separated_dash_rule_is_not_an_empty_setext_heading(
    newline: str,
    trailing: str,
) -> None:
    source = newline.join(("A", "", f"---{trailing}"))

    assert handle_setext_headings(source) == source


@pytest.mark.parametrize("newline", ["\n", "\r\n"])
@pytest.mark.parametrize("trailing", ["", "   ", "\t"])
def test_genuine_setext_h2_accepts_line_local_trailing_space(
    newline: str,
    trailing: str,
) -> None:
    source = newline.join(("A title", f"---{trailing}", "Body"))

    assert handle_setext_headings(source) == newline.join(("## A title", "Body"))


@pytest.mark.parametrize("underline", ["---", "==="])
def test_whitespace_only_predecessor_is_not_a_setext_heading(underline: str) -> None:
    source = f"A\r\n \t\r\n{underline}\r\nB"

    assert handle_setext_headings(source) == source


@pytest.mark.parametrize("fence", ["```", "````", "~~~"])
def test_setext_like_lines_inside_fenced_code_are_literal(fence: str) -> None:
    source = f"{fence}markdown\n---\nReportName: example\nUnit: office\n---\n{fence}\nReal heading\n---\n"

    expected = f"{fence}markdown\n---\nReportName: example\nUnit: office\n---\n{fence}\n## Real heading\n"
    assert handle_setext_headings(source) == expected


def test_inline_code_is_not_rewritten_as_html_break() -> None:
    source = "Text <br> next and `<br>` literal."

    assert normalize_html_tags(source) == "Text   \n next and `<br>` literal."


def test_fenced_html_example_is_not_rewritten() -> None:
    source = "```html\n<br>\n```\nOutside<br>next"

    assert normalize_html_tags(source) == "```html\n<br>\n```\nOutside  \nnext"
