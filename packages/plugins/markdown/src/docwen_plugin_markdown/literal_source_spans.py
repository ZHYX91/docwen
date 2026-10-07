"""Context-aware comment and URL ownership on a protected source projection."""

from __future__ import annotations

import re

_COMMENT_OR_URL_RE = re.compile(r"<!--|%%|https?://[^\s<]+")


def comment_and_url_spans(source: str, *, semantic_url_suffix: bool = False) -> list[tuple[int, int]]:
    """Find literal spans after renderer atoms have retained their delimiters.

    Scan in source order: a URL owns comment-looking characters only outside
    a comment. Once a comment opens, its closer belongs to that comment even
    when it immediately follows a URL. Callers pass a same-length projection
    with code, math and other literal metadata already protected.
    """

    spans: list[tuple[int, int]] = []
    cursor = 0
    while match := _COMMENT_OR_URL_RE.search(source, cursor):
        token = match.group(0)
        if token in {"<!--", "%%"}:
            closer = "-->" if token == "<!--" else "%%"
            closing_start = source.find(closer, match.end())
            end = len(source) if closing_start < 0 else closing_start + len(closer)
        else:
            suffix = token.find("@[[") if semantic_url_suffix else -1
            end = match.end() if suffix < 0 else match.start() + suffix
        spans.append((match.start(), end))
        cursor = end
    return spans
