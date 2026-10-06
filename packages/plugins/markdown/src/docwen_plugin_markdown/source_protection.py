"""Shared source ranges where Markdown-looking semantic tokens stay literal."""

from __future__ import annotations

import re

_BARE_URL_RE = re.compile(r"https?://(?:(?!@\[\[|\[\^)[^\s<])+")
_HTML_TAG_RE = re.compile(r"<[^>\r\n]*>")


def is_markdown_escaped(source: str, start: int) -> bool:
    """Return True when the token at start has an odd backslash prefix."""

    slashes = 0
    cursor = start - 1
    while cursor >= 0 and source[cursor] == "\\":
        slashes += 1
        cursor -= 1
    return slashes % 2 == 1


def markdown_semantic_protected_ranges(source: str) -> list[tuple[int, int]]:
    """Return literal ranges semantic scanners must not reinterpret."""

    ranges: list[tuple[int, int]] = []
    ranges.extend(_paired_ranges(source, "%%", "%%"))
    ranges.extend(_paired_ranges(source, "<!--", "-->"))
    ranges.extend(_escaped_semantic_token_ranges(source))
    ranges.extend(_inline_code_ranges(source))
    ranges.extend(_inline_link_destination_ranges(source))
    ranges.extend(match.span() for match in _HTML_TAG_RE.finditer(source))
    ranges.extend(match.span() for match in _BARE_URL_RE.finditer(source))
    ranges.extend(_indented_code_line_ranges(source))
    return _merge_ranges(ranges)


def overlaps_protected_range(
    protected: list[tuple[int, int]],
    start: int,
    end: int,
) -> bool:
    return any(start < range_end and range_start < end for range_start, range_end in protected)


def _paired_ranges(source: str, opener: str, closer: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    cursor = 0
    while cursor < len(source):
        start = source.find(opener, cursor)
        if start < 0:
            break
        close = source.find(closer, start + len(opener))
        if close < 0:
            ranges.append((start, len(source)))
            break
        end = close + len(closer)
        ranges.append((start, end))
        cursor = end
    return ranges


def _escaped_semantic_token_ranges(source: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    token_patterns = (
        re.compile(r"@\[\[[^\]\r\n]+\]\]"),
        re.compile(r"\[\[[^\]\r\n]+\]\]"),
        re.compile(r"\[\^[^\]\r\n]+\]"),
    )
    for pattern in token_patterns:
        for match in pattern.finditer(source):
            if is_markdown_escaped(source, match.start()):
                ranges.append(match.span())
    return ranges


def _inline_code_ranges(source: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    cursor = 0
    tick_char = chr(96)
    while cursor < len(source):
        tick = source.find(tick_char, cursor)
        if tick < 0:
            break
        run_end = tick
        while run_end < len(source) and source[run_end] == tick_char:
            run_end += 1
        delimiter = source[tick:run_end]
        close = source.find(delimiter, run_end)
        if close < 0:
            cursor = run_end
            continue
        end = close + len(delimiter)
        ranges.append((tick, end))
        cursor = end
    return ranges


def _inline_link_destination_ranges(source: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    cursor = 0
    while cursor < len(source) - 1:
        opening = source.find("](", cursor)
        if opening < 0:
            break
        start = opening + 2
        depth = 1
        index = start
        while index < len(source):
            char = source[index]
            if char == "\\":
                index += 2
                continue
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    ranges.append((start, index))
                    cursor = index + 1
                    break
            if char in "\r\n" and depth == 1:
                cursor = index + 1
                break
            index += 1
        else:
            cursor = start
    return ranges


def _indented_code_line_ranges(source: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    offset = 0
    for raw in source.splitlines(keepends=True):
        text = raw.rstrip("\r\n")
        if text.startswith("    ") or text.startswith("\t"):
            ranges.append((offset, offset + len(text)))
        offset += len(raw)
    return ranges


def _merge_ranges(ranges: list[tuple[int, int]]) -> list[tuple[int, int]]:
    ordered = sorted((start, end) for start, end in ranges if end > start)
    merged: list[tuple[int, int]] = []
    for start, end in ordered:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


__all__ = [
    "is_markdown_escaped",
    "markdown_semantic_protected_ranges",
    "overlaps_protected_range",
]
