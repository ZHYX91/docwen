"""Encode source prose before adding generated Markdown syntax."""

from __future__ import annotations

import re

_INLINE = re.compile(r"([\\`*_{}\[\]<>!#|&~^=$%@])")
_CAPTION_PREFIX = re.compile(r"(?im)^([ \t]*(?:Figure|Table|Equation|Code)):")
_BLOCK_PREFIX = re.compile(r"(?m)^([ \t]*)([-+]|[0-9]+[.)])(?=[ \t])")


def escape_markdown_literal(text: str) -> str:
    """Preserve literal prose, including entities, across inline/block parsers.

    This operates on a source fragment, never on generated Markdown or code.
    Equal-policy runs must be joined before encoding possible block prefixes.
    """
    escaped = _INLINE.sub(r"\\\1", text)
    escaped = _CAPTION_PREFIX.sub(lambda match: match[1] + "\\:", escaped)
    escaped = re.sub(r"(?m)^([ \t]*)(-{3,})([ \t]*)$", lambda match: match[1] + "\\" + match[2] + match[3], escaped)
    return _BLOCK_PREFIX.sub(lambda match: match[1] + match[2][:-1] + "\\" + match[2][-1], escaped)
