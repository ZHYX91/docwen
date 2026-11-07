"""Plain link labels retain content while removing supported formatting."""

from __future__ import annotations

import pytest

from docwen_plugin_markdown.renderer_inlines import plain_link_label

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("Energy $mc^2$ result", "Energy $mc^2$ result"),
        ("$mc^2$", "$mc^2$"),
        ("$$mc^2$$", "$$mc^2$$"),
        (r"**Value $\alpha + \beta$** and `a_b`", r"Value $\alpha + \beta$ and a_b"),
        ("# Heading", "# Heading"),
        ("1. Item", "1. Item"),
        ("- Item", "- Item"),
        ("<span>source</span> and **bold**", "<span>source</span> and bold"),
        ("before<br>after", "before after"),
        ("first\nsecond", "first second"),
        ("![Alt $x$](https://example.com/image.png)", "Alt $x$"),
        ("`$literal$`", "$literal$"),
        ("note [^unbound]", "note [^unbound]"),
        (r"\*literal\*", "*literal*"),
        ("\ufeff$mc^2$", "\ufeff$mc^2$"),
    ],
)
def test_plain_link_label_preserves_inline_content(source: str, expected: str) -> None:
    assert plain_link_label(source) == expected
