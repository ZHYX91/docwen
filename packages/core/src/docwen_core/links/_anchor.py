"""Anchor resolution and content extraction for Markdown links.

Covers:
- parse_anchor: split a link target into (file_path, heading, block_id)
- extract_section_by_heading: extract a heading-delimited section
- extract_block_by_id: extract a paragraph by ``^block-id`` marker
- strip_yaml_front_matter: remove YAML front matter from Markdown content
"""

from __future__ import annotations

import logging
import re
from urllib.parse import unquote

from docwen_core.links._block_context import (
    THEMATIC_BREAK_RE,
    MarkdownBlockContext,
    _Container,
    _indent_end,
    markdown_table_blocks,
    yaml_front_matter_end,
)
from docwen_core.links._source_document import (
    markdown_heading_ranges,
    markdown_heading_text,
    markdown_source_context_lines,
    markdown_source_projection,
)

logger = logging.getLogger(__name__)

_BLOCK_ID = re.compile(r"[A-Za-z0-9-]{1,128}\Z")
_BLOCK_SUFFIX = re.compile(r"(?<!\S)\^([A-Za-z0-9-]{1,128})[ \t]*(?:\r\n|\r|\n)?$")


def parse_anchor(link_target: str) -> tuple[str | None, str | None, str | None]:
    """Parse a link target into (file_path, heading, block_id).

    Supports the following anchor forms commonly used in Obsidian /
    wiki-style links:

    - ``file.md``                     → (file.md, None, None)
    - ``file.md#heading``            → (file.md, heading, None)
    - ``file.md#^block-id``          → (file.md, None, block-id)
    - Structural query/fragment delimiters are split before URL-decoding so
      encoded literal delimiters remain part of the file path.
    - Query strings (``?...``) are stripped.
    """
    path_and_query, separator, raw_anchor = link_target.partition("#")
    raw_path = path_and_query.split("?", 1)[0]
    file_path = unquote(raw_path).strip() or None
    anchor = unquote(raw_anchor).strip() if separator else None

    if not anchor:
        return (file_path, None, None)

    # Block-id anchor (Obsidian-style ``^abc123``)
    if anchor.startswith("^"):
        return (file_path, None, anchor[1:] or None)

    # Heading anchor
    return (file_path, anchor, None)


def extract_section_by_heading(
    content: str, heading: str, *, structural_tables: bool = False, extended_headings: bool = True
) -> str | None:
    """Extract a section of Markdown content bounded by a heading.

    Finds the line whose heading text matches *heading* (case-insensitive,
    whitespace-normalised) and returns everything from that heading line
    through to the next heading of the same or higher level (i.e. same or
    fewer ``#``).  Trailing blank lines at the end of the section are
    trimmed.

    Returns *None* when the heading is not found. Direct callers retain the
    historical H1..H9 default; conversion routes pass their frozen dialect.
    """
    headings = markdown_heading_ranges(
        content, structural_tables=structural_tables, extended_headings=extended_headings
    )
    target_text = "".join(heading.lower().split())
    matches = [
        item
        for item in headings
        if "".join(
            markdown_heading_text("\n".join(content[start:end] for start, end in item.text_ranges)).lower().split()
        )
        == target_text
    ]
    if not matches:
        logger.warning("Requested heading was not found")
        return None
    target = matches[0]
    end = next(
        (item.start for item in headings if item.start > target.start and item.level <= target.level), len(content)
    )
    section = content[target.start : end]
    if target.container_indent >= 4:
        section = "".join(
            line[(_indent_end(line, 0, target.container_indent) or 0) :] for line in section.splitlines(keepends=True)
        )
    return section.rstrip(" \t\r\n")


def extract_block_by_id(
    content: str, block_id: str, *, structural_tables: bool = False, extended_headings: bool = True
) -> str | None:
    """Extract a paragraph marked with a ``^block-id`` marker.

    Block markers come in two flavours:

    **Standalone marker** — ``^block-id`` on its own line.
    Returns the preceding non-blank paragraph.

    **Inline marker** — ``some text ^block-id`` at the end of a line.
    The marker is stripped and adjacent non-heading lines above and below
    are gathered into the result.
    """
    if _BLOCK_ID.fullmatch(block_id) is None:
        return None
    projection = markdown_source_projection(
        content, structural_tables=structural_tables, extended_headings=extended_headings
    )
    lines = content.splitlines(keepends=True)
    visible_lines = projection.splitlines(keepends=True)
    offsets: list[int] = []
    owners: list[tuple[_Container, ...]] = []
    content_starts: list[int] = []
    contexts = markdown_source_context_lines(
        content, structural_tables=structural_tables, extended_headings=extended_headings
    )
    offset = 0
    matches: list[tuple[int, re.Match[str]]] = []
    valid_markers: dict[int, re.Match[str]] = {}
    for index, line in enumerate(visible_lines):
        row = contexts[index]
        offsets.append(offset)
        owners.append(row.containers)
        content_starts.append(row.content_start)
        match = _BLOCK_SUFFIX.search(line)
        original_match = _BLOCK_SUFFIX.search(lines[index])
        if (
            match is not None
            and original_match is not None
            and match.start() == original_match.start()
            and match.group(1) == original_match.group(1)
        ):
            valid_markers[index] = original_match
            if original_match.group(1) == block_id:
                matches.append((index, original_match))
        offset += len(line)
    if len(matches) != 1:
        logger.warning("Requested block id was not found or is ambiguous")
        return None
    index, match = matches[0]
    heading_ranges = markdown_heading_ranges(
        content, structural_tables=structural_tables, extended_headings=extended_headings
    )
    table_ranges = [
        (start, end - 1)
        for start, end in markdown_table_blocks(
            [line[content_starts[index] :].rstrip("\r\n") for index, line in enumerate(visible_lines)],
            list(owners),
            structural_tables=structural_tables,
            extended_headings=extended_headings,
        )
    ]

    def ordinary_line(candidate: int, *, allow_marker: bool = False) -> bool:
        if not 0 <= candidate < len(lines) or not visible_lines[candidate].strip():
            return False
        if owners[candidate] != owners[index]:
            return False
        if not allow_marker and candidate in valid_markers:
            return False
        body = visible_lines[candidate][content_starts[candidate] :].rstrip("\r\n")
        if THEMATIC_BREAK_RE.fullmatch(body) or any(start <= candidate <= end for start, end in table_ranges):
            return False
        return not any(item.start <= offsets[candidate] < item.end for item in heading_ranges)

    standalone = visible_lines[index][content_starts[index] :].strip() == f"^{block_id}"
    if standalone:
        end_index = index - 1
        while end_index >= 0 and not lines[end_index].strip():
            end_index -= 1
        if end_index < 0:
            return None
        parent = owners[index]
        previous = owners[end_index]
        if len(previous) > len(parent) and previous[: len(parent)] == parent:
            container = previous[len(parent)]
            start_index = end_index
            for candidate in range(end_index - 1, -1, -1):
                path = owners[candidate]
                if not lines[candidate].strip():
                    start_index = candidate
                    continue
                if len(path) <= len(parent) or path[: len(parent)] != parent:
                    break
                sibling = path[len(parent)]
                if sibling != container and not (
                    container.kind == sibling.kind == "list" and container.marker == sibling.marker
                ):
                    break
                start_index = candidate
            return "".join(lines[start_index : end_index + 1]).strip("\r\n")
        for start, end in table_ranges:
            if end == end_index and owners[start] == owners[index]:
                return "".join(lines[start : end + 1]).rstrip("\r\n")
        if not ordinary_line(end_index):
            return _extract_preceding_literal_block(
                content,
                offsets[end_index] + len(lines[end_index]),
                structural_tables=structural_tables,
                extended_headings=extended_headings,
            )
        start_index = end_index
        while ordinary_line(start_index - 1):
            start_index -= 1
        return "".join(lines[start_index : end_index + 1]).rstrip(" \t\r\n")

    if not ordinary_line(index, allow_marker=True):
        return None
    start_index = index
    if owners[index] and owners[index][-1].kind == "list":
        item = owners[index]
        while start_index > 0 and owners[start_index - 1][: len(item)] == item:
            start_index -= 1
        end_index = index
        while end_index + 1 < len(lines) and owners[end_index + 1][: len(item)] == item:
            end_index += 1
        clean_line = lines[index][: match.start()].rstrip(" \t") + lines[index][len(lines[index].rstrip("\r\n")) :]
        return ("".join(lines[start_index:index]) + clean_line + "".join(lines[index + 1 : end_index + 1])).rstrip(
            " \t\r\n"
        )
    while ordinary_line(start_index - 1):
        start_index -= 1
    end_index = index
    while ordinary_line(end_index + 1):
        end_index += 1
    clean_line = lines[index][: match.start()].rstrip(" \t") + lines[index][len(lines[index].rstrip("\r\n")) :]
    return ("".join(lines[start_index:index]) + clean_line + "".join(lines[index + 1 : end_index + 1])).rstrip(
        " \t\r\n"
    )


def _extract_preceding_literal_block(
    content: str, end: int, *, structural_tables: bool = False, extended_headings: bool = False
) -> str | None:
    from docwen_core.links._source_ownership import markdown_source_owners

    for owner in markdown_source_owners(
        content, full_document=True, structural_tables=structural_tables, extended_headings=extended_headings
    ):
        if owner.kind == "block" and owner.start < end <= owner.end and not content[end : owner.end].strip():
            segment = content[owner.start : owner.end]
            context = MarkdownBlockContext(extended_headings=extended_headings)
            first = context.line(segment.splitlines()[0]).content
            if re.match(r" {0,3}(?:`{3,}|~{3,}|\$\$)", first):
                return segment.rstrip(" \t\r\n")
            return None
    return None


def split_yaml_front_matter_source(content: str) -> tuple[str, str]:
    """Return an exact YAML front-matter prefix and the remaining body."""
    end = yaml_front_matter_end(content)
    return content[:end], content[end:]


def strip_yaml_front_matter(content: str) -> str:
    """Remove YAML front matter (``---`` delimited) from the start of
    Markdown content.

    Returns the content unchanged when no YAML front matter is detected.
    """
    yaml_front, body = split_yaml_front_matter_source(content)
    if not yaml_front:
        return content
    result = body.lstrip("\r\n")
    logger.debug("Stripped YAML front matter")
    return result
