"""Extended mistune Markdown parser with custom inline plugins.

Uses mistune v3's plugin API (``md.inline.register``). The built-in
formatting plugins from ``mistune.plugins.formatting`` provide:

- ``==highlight==``  — via plugin ``mark`` (``<mark>`` element)
- ``^^insert^^``    — via plugin ``insert`` (``<ins>`` element)
- ``^superscript^`` — via plugin ``superscript``
- ``~subscript~``   — via plugin ``subscript``

We also provide a custom ``plugin_underline`` for ``<u>text</u>`` inline
HTML support (rendered as underline in AST mode).

The module exposes ``parse_markdown_text()`` as the Markdown plugin's parser
entry point.
"""

from __future__ import annotations

import re
from typing import Any

import mistune
from mistune.plugins.formatting import (
    insert as _insert,
)
from mistune.plugins.formatting import (
    mark as _mark,
)
from mistune.plugins.formatting import (
    subscript as _subscript,
)
from mistune.plugins.formatting import (
    superscript as _superscript,
)

_EXTENDED_ATX_HEADING_TRIM = re.compile(r"(\s+|^)#+\s*$")

# ═══════════════════════════════════════════════════════════════════════════
# Custom plugin: <u>underline</u>
# ═══════════════════════════════════════════════════════════════════════════


def plugin_underline(md: mistune.Markdown) -> None:
    """Mistune v3 plugin: inline ``<u>text</u>`` → underline token.

    Registers a parse rule that captures ``<u>...</u>`` and emits a
    ``underline`` token in AST mode.
    """

    PATTERN = r"<u>(.*?)</u>"

    def parse_underline(inline, m, state):
        text = m.group(1)
        new_state = state.copy()
        new_state.src = text
        children = inline.render(new_state)
        state.append_token({"type": "underline", "children": children})
        return m.end()

    md.inline.register(
        "underline",
        PATTERN,
        parse_underline,
        before="linebreak",
    )


# ═══════════════════════════════════════════════════════════════════════════
# Custom plugin: standalone single-line display math
# ═══════════════════════════════════════════════════════════════════════════


def plugin_single_line_block_math(md: mistune.Markdown) -> None:
    """Parse a standalone ``$$formula$$`` line as ``block_math``.

    Mistune's built-in math plugin only treats the three-line form as display
    math.  The maintained parser contract also accepts the common single-line form; if
    it falls through to the inline parser, the two outer dollar signs become
    literal text in Word.  Registering this as a block rule keeps fenced code
    untouched and gives both spellings the same AST contract.
    """

    pattern = r"^ {0,3}\$\$(?P<single_line_math_text>[^\n]+?)\$\$[ \t]*$"

    def parse_single_line_block_math(block, match, state):
        state.append_token(
            {
                "type": "block_math",
                "raw": match.group("single_line_math_text").strip(),
            }
        )
        return match.end() + 1

    md.block.register(
        "single_line_block_math",
        pattern,
        parse_single_line_block_math,
        before="block_math",
    )


def plugin_extended_atx_headings(md: mistune.Markdown) -> None:
    """Parse DocWen's legacy-compatible ATX heading levels 7 through 9.

    CommonMark and Mistune stop at level 6.  Word exposes nine outline
    levels, and classic DocWen represented the last three with seven to nine
    leading ``#`` markers.  Keep that extension narrow: ten or more markers
    remain ordinary paragraph text.
    """

    pattern = (
        r"^ {0,3}(?P<extended_atx_marks>#{7,9})(?!#)"
        r"(?P<extended_atx_text>[ \t]*|[ \t]+.*?)$"
    )

    def parse_extended_atx_heading(block, match, state):
        text = match.group("extended_atx_text").strip()
        if text:
            text = _EXTENDED_ATX_HEADING_TRIM.sub("", text)
        state.append_token(
            {
                "type": "heading",
                "text": text,
                "attrs": {"level": len(match.group("extended_atx_marks"))},
                "style": "atx",
            }
        )
        return match.end() + 1

    md.block.register(
        "extended_atx_heading",
        pattern,
        parse_extended_atx_heading,
        before="atx_heading",
    )


# ═══════════════════════════════════════════════════════════════════════════
# Factory
# ═══════════════════════════════════════════════════════════════════════════


def create_extended_markdown(*, auto_link_bare_url: bool = False) -> mistune.Markdown:
    """Create a mistune Markdown instance with the requested plugins.

    Includes:
    - Built-in: table, footnotes, strikethrough, task_lists, math
    - Formatting: mark (``==text==``), insert (``^^text^^``),
      superscript (``^text^``), subscript (``~text~``)
    - Custom: underline (``<u>text</u>``)

    Returns:
        A mistune ``Markdown`` instance configured with ``renderer="ast"``
        for structured output.
    """
    plugins: list[Any] = [
        "table",
        "footnotes",
        "strikethrough",
        "task_lists",
        "math",
        _mark,
        _insert,
        _superscript,
        _subscript,
        plugin_single_line_block_math,
        plugin_extended_atx_headings,
        plugin_underline,
    ]
    if auto_link_bare_url:
        plugins.append("url")

    return mistune.create_markdown(
        renderer="ast",
        plugins=plugins,
    )


# ═══════════════════════════════════════════════════════════════════════════
# Parsing entry point
# ═══════════════════════════════════════════════════════════════════════════


def parse_markdown_text(
    content: str,
    *,
    auto_link_bare_url: bool = False,
) -> list[dict[str, Any]]:
    """Parse Markdown text into an AST using the extended mistune parser.

    Args:
        content: Raw Markdown text.

    Returns:
        A list of token dicts from mistune's AST renderer.
    """
    parser = create_extended_markdown(auto_link_bare_url=auto_link_bare_url)
    return parser(content)  # pyright: ignore[reportReturnType]
