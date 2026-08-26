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

logger = logging.getLogger(__name__)

_YAML_FRONT_OPEN_RE = re.compile(r"^(?:\ufeff)?---[ \t]*(?:\r?\n)")
_YAML_FRONT_CLOSE_RE = re.compile(r"^---[ \t]*(?:\r?\n|$)", re.MULTILINE)


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


def extract_section_by_heading(content: str, heading: str) -> str | None:
    """Extract a section of Markdown content bounded by a heading.

    Finds the line whose heading text matches *heading* (case-insensitive,
    whitespace-normalised) and returns everything from that heading line
    through to the next heading of the same or higher level (i.e. same or
    fewer ``#``).  Trailing blank lines at the end of the section are
    trimmed.

    Returns *None* when the heading is not found.
    """
    lines = content.split("\n")
    target_heading = heading.strip()

    start_index: int | None = None
    start_level: int | None = None

    for i, line in enumerate(lines):
        m = re.match(r"^(#{1,9})\s+(.*)$", line)
        if not m:
            continue
        level = len(m.group(1))
        title = m.group(2).strip()
        if title.lower().replace(" ", "") == target_heading.lower().replace(" ", ""):
            start_index = i
            start_level = level
            logger.debug("Found target heading: '%s' (level %d) at line %d", title, level, i)
            break

    if start_index is None:
        logger.warning("Heading not found: '%s'", heading)
        return None
    if start_level is None:
        return None

    # Find end of section — next heading at same or higher level
    end_index = len(lines)
    for i in range(start_index + 1, len(lines)):
        m = re.match(r"^(#{1,9})\s+", lines[i])
        if m and len(m.group(1)) <= start_level:
            end_index = i
            logger.debug("Section ends at line %d (heading level %s)", i, m.group(1))
            break

    section_lines = lines[start_index:end_index]

    # Trim trailing blank lines
    while section_lines and not section_lines[-1].strip():
        section_lines.pop()

    result = "\n".join(section_lines)
    logger.debug("Extracted section: %d lines", len(section_lines))
    return result


def extract_block_by_id(content: str, block_id: str) -> str | None:
    """Extract a paragraph marked with a ``^block-id`` marker.

    Block markers come in two flavours:

    **Standalone marker** — ``^block-id`` on its own line.
    Returns the preceding non-blank paragraph.

    **Inline marker** — ``some text ^block-id`` at the end of a line.
    The marker is stripped and adjacent non-heading lines above and below
    are gathered into the result.
    """
    lines = content.split("\n")
    block_re = re.compile(r"\^" + re.escape(block_id) + r"\s*$")

    for i, line in enumerate(lines):
        if not block_re.search(line):
            continue

        line_stripped = line.strip()

        # --- standalone marker: ^block-id on its own line ---
        if line_stripped == f"^{block_id}":
            logger.debug("Found standalone block marker at line %d", i)
            paragraph_lines: list[str] = []
            j = i - 1
            # skip blank lines above the marker
            while j >= 0 and not lines[j].strip():
                j -= 1
            # collect non-blank lines (the paragraph)
            while j >= 0 and lines[j].strip():
                paragraph_lines.insert(0, lines[j])
                j -= 1
            if paragraph_lines:
                result = "\n".join(paragraph_lines)
                logger.debug("Extracted block (standalone): %d lines", len(paragraph_lines))
                return result
            continue

        # --- inline marker: text ^block-id in a line ---
        clean_line = block_re.sub("", line).rstrip()
        paragraph_lines = [clean_line]

        # scan upwards for adjacent non-heading lines
        j = i - 1
        while j >= 0 and lines[j].strip() and not lines[j].strip().startswith("#"):
            paragraph_lines.insert(0, lines[j])
            j -= 1

        # scan downwards for adjacent non-heading lines
        j = i + 1
        while j < len(lines) and lines[j].strip() and not lines[j].strip().startswith("#"):
            # stop if we encounter another block marker
            if re.search(r"\^\w+", lines[j]):
                break
            paragraph_lines.append(lines[j])
            j += 1

        result = "\n".join(paragraph_lines)
        logger.debug("Extracted block (inline): %d lines", len(paragraph_lines))
        return result

    logger.warning("Block id not found: ^%s", block_id)
    return None


def split_yaml_front_matter_source(content: str) -> tuple[str, str]:
    """Return an exact YAML front-matter prefix and the remaining body."""
    opening = _YAML_FRONT_OPEN_RE.match(content)
    if opening is None:
        return "", content
    closing = _YAML_FRONT_CLOSE_RE.search(content, opening.end())
    if closing is None:
        return "", content
    return content[: closing.end()], content[closing.end() :]


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
