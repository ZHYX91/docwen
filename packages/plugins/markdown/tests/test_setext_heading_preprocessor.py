"""Setext preprocessing preserves blank-separated thematic breaks."""

from __future__ import annotations

import pytest

from docwen_plugin_markdown.preprocessor import handle_setext_headings

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
