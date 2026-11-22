"""Reference definitions and image occurrences in the authored Markdown grammar."""

from __future__ import annotations

import html
import re
import string
from bisect import bisect_left, bisect_right
from collections.abc import Mapping
from contextvars import ContextVar
from dataclasses import dataclass
from urllib.parse import quote

from docwen_core.links._markdown_inline import (
    MarkdownInlineLink,
    _contains_active_protected_token,
    _has_blank_line,
    _is_backslash_escaped,
    _is_markdown_blank_line,
    _is_valid_footnote_definition_label,
    _label_starts_a_new_block,
    _link_label_close,
    _normalize_reference_label,
    _reference_definition_parts,
    parse_inline_link,
)

_LABEL = r"(?:\r\n|[^\r\\\[\]]|\r(?!\n)|\\[^\r\n]){0,500}"
_DEFINITION = re.compile(r" {0,3}\[(?P<label>" + _LABEL + r")\]:")
_TAIL = re.compile(r"\[(?P<label>" + _LABEL + r")\]")
_LINE_END = re.compile(r"[ \t]*(?:\r\n|\r|\n|$)")
_EOL = r"(?:\r\n|\r(?!\n)|(?<!\r)\n)"
_BLANK_LINE = re.compile(_EOL + r"[ \t\v\f]*" + _EOL)
_SPACE = " \t\r\n\f"
_DESTINATION_ESCAPE = re.compile(
    r"\\([" + re.escape(string.punctuation) + r"])|&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]+|[^\t\n\f <&#;]{1,32});"
)


@dataclass(frozen=True, slots=True)
class MarkdownReferenceDefinition:
    label: str
    target: str
    end: int


@dataclass(slots=True)
class FootnoteDefinitionContext:
    """Keep the renderer's footnote body out of global reference definitions."""

    containers: tuple[object, ...] = ()
    pending_body_line: bool = False
    continuation_lead: int | None = None

    def consumes(self, content: str, containers: tuple[object, ...]) -> bool:
        line = content.rstrip("\r\n")
        blank = _is_markdown_blank_line(line)
        if containers != self.containers:
            self.pending_body_line = False
            self.continuation_lead = None
        self.containers = containers
        if self.pending_body_line:
            self.pending_body_line = False
            if not blank:
                return True
        if self.continuation_lead is not None:
            continuation = re.fullmatch(rf" {{{self.continuation_lead + 1},{self.continuation_lead + 4}}}(?! ).*", line)
            if not blank and continuation is not None:
                return True
            if not blank:
                self.continuation_lead = None
        definition = _reference_definition_parts(line)
        if definition is None:
            return False
        label, payload = definition
        if not _is_valid_footnote_definition_label(label) or (payload and payload[0] not in " \t"):
            return False
        self.pending_body_line = payload == ""
        self.continuation_lead = len(line) - len(line.lstrip(" "))
        return True


_ACTIVE_REFERENCE_DEFINITIONS: ContextVar[Mapping[str, MarkdownReferenceDefinition] | None] = ContextVar(
    "docwen_reference_definitions", default=None
)


@dataclass(frozen=True, slots=True)
class ReferenceImageSource:
    """Container-neutral recognition with exact authored token coordinates."""

    source: str
    neutral: str
    prefixes: tuple[tuple[int, int], ...]
    prefix_starts: tuple[int, ...]
    boundaries: tuple[int, ...]

    def contains(self, start: int, end: int) -> bool:
        index = bisect_right(self.boundaries, start)
        return index < len(self.boundaries) and end <= self.boundaries[index]

    def label(self, start: int, end: int) -> str:
        parts: list[str] = []
        cursor = start
        index = bisect_right(self.prefix_starts, start)
        if index and self.prefixes[index - 1][1] > start:
            index -= 1
        for offset in range(index, len(self.prefixes)):
            prefix_start, prefix_end = self.prefixes[offset]
            if prefix_start >= end:
                break
            parts.append(self.source[cursor:prefix_start])
            cursor = min(prefix_end, end)
        parts.append(self.source[cursor:end])
        return "".join(parts)

    def region(self, start: int, end: int) -> ReferenceImageSource:
        first = bisect_left(self.prefix_starts, start)
        last = bisect_left(self.prefix_starts, end)
        prefixes = tuple((left - start, right - start) for left, right in self.prefixes[first:last])
        first_boundary = bisect_right(self.boundaries, start)
        last_boundary = bisect_left(self.boundaries, end)
        return ReferenceImageSource(
            self.source[start:end],
            self.neutral[start:end],
            prefixes,
            tuple(left for left, _right in prefixes),
            (0, *(boundary - start for boundary in self.boundaries[first_boundary:last_boundary]), end - start),
        )


_ACTIVE_REFERENCE_IMAGE_SOURCE: ContextVar[ReferenceImageSource | None] = ContextVar(
    "docwen_reference_image_source", default=None
)


def reference_image_source(
    source: str, *, structural_tables: bool = False, extended_headings: bool = False
) -> ReferenceImageSource:
    """Blank only actual prefixes; inline scopes prohibit crossing container exits."""
    from docwen_core.links._block_context import MarkdownBlockContext, markdown_inline_regions, markdown_source_lines

    context = MarkdownBlockContext(extended_headings=extended_headings)
    chars = list(source)
    prefixes: list[tuple[int, int]] = []
    offset = 0
    for line in markdown_source_lines(source):
        row = context.line(line, allow_lazy=True)
        if row.content_start:
            end = offset + row.content_start
            prefixes.append((offset, end))
            chars[offset:end] = " " * row.content_start
        offset += len(line)
    regions = markdown_inline_regions(source, structural_tables=structural_tables, extended_headings=extended_headings)
    return ReferenceImageSource(
        source,
        "".join(chars),
        tuple(prefixes),
        tuple(start for start, _end in prefixes),
        (0, *(end for _start, end in regions)),
    )


def parse_reference_definition(text: str) -> tuple[MarkdownReferenceDefinition | None, str | None]:
    """Parse a logical prefix and return its pending continuation, if any.

    The caller owns paragraph/container eligibility. A title on the next line
    can extend a complete destination; an invalid title leaves that next line
    in the document. A pending title carries its closer; an empty string asks
    for ordinary label/destination lookahead. None means the parse is complete.
    Offsets always count the original line endings.
    """
    opening = _DEFINITION.match(text)
    if opening is None:
        # A label can continue across lines, but has the renderer's 500-unit cap.
        more = re.match(r" {0,3}\[", text) and "]:" not in text and len(text) <= 1004
        return None, "" if more else None
    label = opening.group("label")
    if not _normalize_reference_label(label):
        return None, None
    payload = text[opening.end() :]
    if _is_valid_footnote_definition_label(label) and (not payload or payload[0] in " \t\r\n"):
        return None, None
    position = opening.end()
    whitespace = re.match(r"[ \t]*(?:(?:\r\n|\r|\n)[ \t]*)?", text[position:])
    assert whitespace is not None
    position += whitespace.end()
    if position == len(text):
        return None, "" if not _has_blank_line(text) else None
    if text[position] in _SPACE:
        return None, None
    target_start = position
    if text[position] == "<":
        position += 1
        while position < len(text) and text[position] != ">":
            if text[position] in "<\\\r\n\x00":
                return None, None
            position += 1
        if position == len(text):
            return None, None
        position += 1
    else:
        while position < len(text) and text[position] not in _SPACE:
            if text[position] == "\x00":
                return None, None
            position += 2 if text[position] == "\\" and position + 1 < len(text) else 1
    target = text[target_start:position]
    fallback_end = _LINE_END.match(text, position)
    fallback = MarkdownReferenceDefinition(label, target, fallback_end.end()) if fallback_end else None
    title_start = position
    while title_start < len(text) and text[title_start] in _SPACE:
        title_start += 1
    if title_start == len(text):
        return fallback, ""
    if title_start == position or _BLANK_LINE.search(text[position:title_start]):
        return fallback, None
    closer = {"'": "'", '"': '"', "(": ")"}.get(text[title_start])
    if closer is None:
        return fallback, None
    cursor = title_start + 1
    while cursor < len(text):
        if text[cursor] == "\x00" or _BLANK_LINE.match(text, cursor):
            return fallback, None
        if text[cursor] == "\\":
            cursor += 2
            continue
        if text[cursor] == closer:
            end = _LINE_END.match(text, cursor + 1)
            if end is None:
                return fallback, None
            return MarkdownReferenceDefinition(label, target, end.end()), None
        cursor += 1
    return fallback, closer


def markdown_reference_definitions(
    source: str, *, extended_headings: bool = False, include_front_matter: bool = True
) -> dict[str, MarkdownReferenceDefinition]:
    """Collect first definitions from actual visible, eligible source blocks."""
    from docwen_core.links._non_embed import _split_fenced_code_blocks

    definitions: dict[str, MarkdownReferenceDefinition] = {}
    _split_fenced_code_blocks(
        source, extended_headings=extended_headings, references=definitions, include_front_matter=include_front_matter
    )
    return definitions


def _reference_image_target(target: str) -> str:
    """Decode authored escapes/entities once, before the path's URI boundary."""
    raw = target[1:-1] if target.startswith("<") else target

    def decode(match: re.Match[str]) -> str:
        escaped = match.group(1)
        if escaped is not None:
            # Match the inline-image contract for literal URI delimiters.
            return quote(escaped, safe="") if escaped in "#?%<>\\" else escaped
        return html.unescape(match.group(0))

    value = _DESTINATION_ESCAPE.sub(decode, raw)
    return "<" + quote(value, safe=":/?#@!$&()*+,;=%") + ">"


def parse_markdown_image(
    text: str, start: int, definitions: Mapping[str, MarkdownReferenceDefinition] | None = None
) -> MarkdownInlineLink | None:
    """Read inline, full, collapsed or shortcut image syntax at one position."""
    inline = parse_inline_link(text, start, image=True)
    if inline is not None:
        return inline
    if definitions is None:
        definitions = _ACTIVE_REFERENCE_DEFINITIONS.get() or {}
    if not definitions or not text.startswith("![", start) or _is_backslash_escaped(text, start):
        return None
    projection = _ACTIVE_REFERENCE_IMAGE_SOURCE.get()
    if projection is None or projection.source != text:
        projection = reference_image_source(text)
    neutral = projection.neutral
    # An inline destination still wins when container prefixes were the only
    # obstacle to reading the authored image. Source labels retain their length.
    projected_inline = parse_inline_link(neutral, start, image=True)
    if projected_inline is not None and projection.contains(start, projected_inline.end):
        close = start + 2 + len(projected_inline.label)
        return MarkdownInlineLink(
            end=projected_inline.end,
            label=text[start + 2 : close],
            target=projected_inline.target,
            is_image=True,
            display_label=projection.label(start + 2, close),
        )
    close = _link_label_close(neutral, start + 1)
    if close is None:
        return None
    label = neutral[start + 2 : close]
    if _contains_active_protected_token(label) or _has_blank_line(label) or _label_starts_a_new_block(label):
        return None
    end = close + 1
    reference = label
    if end < len(text) and text[end] == "[":
        tail_close = _link_label_close(neutral, end)
        if tail_close is not None and projection.contains(start, tail_close + 1):
            tail = _TAIL.fullmatch(projection.label(end, tail_close + 1))
            if tail is not None:
                reference = tail.group("label") or label
                end = tail_close + 1
    if not projection.contains(start, end):
        return None
    definition = definitions.get(_normalize_reference_label(reference))
    if definition is None:
        return None
    return MarkdownInlineLink(
        end=end,
        label=text[start + 2 : close],
        target=_reference_image_target(definition.target),
        is_image=True,
        display_label=projection.label(start + 2, close),
    )
