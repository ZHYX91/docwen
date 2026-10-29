"""Full-source Markdown ownership and exact heading edit ranges.

The projection keeps every Unicode code point and line ending in place. It is
shared by source numbering, navigation and proofreading, never serialized as
the user's Markdown document.
"""

from __future__ import annotations

import html
import re
from dataclasses import dataclass

from docwen_core.links._block_context import (
    SETEXT_HEADING_RE,
    MarkdownBlockContext,
    MarkdownBlockLine,
    match_atx_heading,
    yaml_front_matter_end,
)
from docwen_core.links._markdown_inline import matching_backtick_run_end, parse_inline_link

_PARAGRAPH_INTERRUPT = re.compile(r" {0,3}(?:>|`{3,}|~{3,}|\$\$|[-+*][ \t]+|\d+[.)][ \t]+)")


def _blank(source: str) -> str:
    return "".join(char if char in "\r\n" else " " for char in source)


def markdown_source_projection(
    source: str,
    *,
    inline_literals: bool = True,
    structural_tables: bool = False,
    extended_headings: bool = False,
) -> str:
    """Hide block/inline literals, comments, YAML and parsed link metadata."""
    from docwen_core.links._source_ownership import markdown_source_owners

    chars = list(source)
    for owner in markdown_source_owners(
        source, full_document=True, structural_tables=structural_tables, extended_headings=extended_headings
    ):
        if not inline_literals and owner.kind not in {"block", "comment"}:
            continue
        chars[owner.start : owner.end] = _blank(source[owner.start : owner.end])
    return "".join(chars)


def markdown_source_context_lines(
    source: str, *, structural_tables: bool = False, extended_headings: bool = False
) -> list[MarkdownBlockLine]:
    """Keep original container prefixes while ignoring literal sample markers."""
    from docwen_core.links._source_ownership import markdown_source_owners

    owners = markdown_source_owners(
        source, full_document=True, structural_tables=structural_tables, extended_headings=extended_headings
    )
    grammar = list(source)
    for owner in owners:
        if owner.kind == "comment":
            grammar[owner.start : owner.end] = _blank(source[owner.start : owner.end])
    front_matter_end = yaml_front_matter_end(source)
    grammar[:front_matter_end] = _blank(source[:front_matter_end])
    blocks = [owner for owner in owners if owner.kind == "block"]
    rows: list[MarkdownBlockLine] = []
    context = MarkdownBlockContext(extended_headings=extended_headings)
    block_index = 0
    offset = 0
    for line in "".join(grammar).splitlines(keepends=True):
        bom = int(offset == 0 and source.startswith("\ufeff"))
        probe = offset + bom
        while block_index < len(blocks) and blocks[block_index].end <= probe:
            block_index += 1
        block = blocks[block_index] if block_index < len(blocks) and blocks[block_index].start <= probe else None
        row = context.line(
            line[bom:],
            allow_new=block is None or block.start >= probe,
            allow_lazy=block is None,
        )
        rows.append(MarkdownBlockLine(row.content, row.content_start + bom, row.containers, row.lazy))
        offset += len(line)
    return rows


@dataclass(frozen=True, slots=True)
class MarkdownHeadingRange:
    start: int
    end: int
    text_start: int
    text_end: int
    level: int
    style: str
    text_ranges: tuple[tuple[int, int], ...] = ()
    container_indent: int = 0


def markdown_heading_ranges(
    source: str, *, extended_headings: bool = True, structural_tables: bool = False
) -> list[MarkdownHeadingRange]:
    """Return actual ATX/Setext heading owners in original source coordinates."""
    projection = markdown_source_projection(
        source, inline_literals=False, structural_tables=structural_tables, extended_headings=extended_headings
    )
    contexts = markdown_source_context_lines(
        source, structural_tables=structural_tables, extended_headings=extended_headings
    )
    headings: list[MarkdownHeadingRange] = []
    paragraph: list[tuple[int, int, int]] = []
    offset = 0
    last_containers: tuple[object, ...] = ()
    for line, block_line in zip(projection.splitlines(keepends=True), contexts, strict=True):
        visible = line[block_line.content_start :].rstrip("\r\n")
        content_start = offset + block_line.content_start
        prefix = source[offset:content_start]
        container_indent = (
            len(prefix.expandtabs(4)) if block_line.containers and prefix and not prefix.strip(" \t") else 0
        )
        if block_line.containers != last_containers:
            paragraph = []
        last_containers = block_line.containers
        atx = match_atx_heading(visible, extended_headings=extended_headings)
        original_line = source[content_start : offset + len(line.rstrip("\r\n"))]
        original_atx = match_atx_heading(original_line, extended_headings=extended_headings)
        if atx is not None and original_atx is not None:
            text_start = content_start + original_atx.end()
            text_end = offset + len(line.rstrip("\r\n"))
            raw_title = source[text_start:text_end]
            closing = re.search(r"[ \t]+#+[ \t]*$", raw_title)
            if closing is not None:
                text_end = text_start + closing.start()
            text_end = text_start + len(source[text_start:text_end].rstrip(" \t"))
            headings.append(
                MarkdownHeadingRange(
                    offset,
                    offset + len(line),
                    text_start,
                    text_end,
                    len(atx.group(1)),
                    "atx",
                    ((text_start, text_end),),
                    container_indent,
                )
            )
            paragraph = []
        elif (
            (underline := SETEXT_HEADING_RE.fullmatch(visible)) is not None
            and SETEXT_HEADING_RE.fullmatch(original_line) is not None
            and not block_line.lazy
            and paragraph
        ):
            start, text_start, _first_end = paragraph[0]
            headings.append(
                MarkdownHeadingRange(
                    start,
                    offset + len(line),
                    text_start,
                    paragraph[-1][2],
                    1 if underline.group(1)[0] == "=" else 2,
                    "setext",
                    tuple((text_start, text_end) for _start, text_start, text_end in paragraph),
                )
            )
            paragraph = []
        elif (
            visible.strip()
            and atx is None
            and _PARAGRAPH_INTERRUPT.match(visible) is None
            and not visible.startswith(("    ", "\t"))
        ):
            paragraph.append(
                (
                    offset,
                    content_start + len(visible) - len(visible.lstrip(" \t")),
                    offset + len(line.rstrip(" \t\r\n")),
                )
            )
        else:
            paragraph = []
        offset += len(line)
    return headings


def markdown_heading_text(source: str) -> str:
    """Plain inline heading text for the documented navigation spellings."""
    parts: list[str] = []
    literals: dict[str, str] = {}

    def literal(value: str) -> None:
        token = f"\x00DOCWEN-HEADING-{len(literals)}\x00"
        while token in source:
            token += "\x00"
        literals[token] = value
        parts.append(token)

    cursor = 0
    while cursor < len(source):
        if (
            source[cursor] == "\\"
            and cursor + 1 < len(source)
            and source[cursor + 1] in "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"
        ):
            literal(source[cursor + 1])
            cursor += 2
            continue
        link = parse_inline_link(source, cursor, image=False) or parse_inline_link(source, cursor, image=True)
        if link is not None:
            literal(markdown_heading_text(link.label))
            cursor = link.end
            continue
        if source[cursor] == "`" and (end := matching_backtick_run_end(source, cursor)) is not None:
            length = len(source[cursor:]) - len(source[cursor:].lstrip("`"))
            value = source[cursor + length : end - length].replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
            if value.startswith(" ") and value.endswith(" ") and value.strip(" "):
                value = value[1:-1]
            literal(value)
            cursor = end
            continue
        parts.append(source[cursor])
        cursor += 1
    text = "".join(parts)
    text = re.sub(r"<(?:br\s*/?|[^>]+)>", " ", text)
    text = re.sub(r"(\*{1,3}|~~|==)(?=\S)(.+?)(?<=\S)\1", r"\2", text, flags=re.DOTALL)
    text = re.sub(r"(?<!\w)(_{1,3})(?=\S)(.+?)(?<=\S)\1(?!\w)", r"\2", text, flags=re.DOTALL)
    text = html.unescape(text)
    for token, value in literals.items():
        text = text.replace(token, value)
    return text
