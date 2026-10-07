"""Common utility functions shared by Markdown conversion routes.

Includes:
- File reading (``read_input_markdown``)
- Heading numbering (``remove_md_numbering``, ``add_md_numbering``)
- Table extraction for spreadsheet conversion
- CSV writing
"""

from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any

from docwen_core.links import restore_table_safe_breaks
from docwen_core.models.file_inspection import FILE_INSPECTION_METADATA_KEY, DetectionMethod
from docwen_core.models.request import ConversionRequest
from docwen_core.text.heading_numbering import (
    HeadingFormatter,
    resolve_heading_numbering_scheme,
    strip_heading_prefix,
)

# ── File reading ───────────────────────────────────────────────────────


def link_source_path(request: ConversionRequest, physical_path: str) -> str:
    """Keep synthetic storage out of the authored-resource lookup boundary."""
    return "" if has_synthetic_source(request) else physical_path


def has_synthetic_source(request: ConversionRequest) -> bool:
    for ref in request.input_refs:
        if ref.input_role != "source":
            continue
        inspection = ref.metadata.get(FILE_INSPECTION_METADATA_KEY)
        if (
            isinstance(inspection, dict)
            and inspection.get("detection_method") == DetectionMethod.SYNTHETIC_MARKDOWN.value
        ):
            return True
    return False


def conversion_error_text(request: ConversionRequest, error: Exception) -> str:
    """Synthetic input diagnostics must not echo authored text from exceptions."""
    return f"Markdown conversion failed ({type(error).__name__})." if has_synthetic_source(request) else str(error)


def read_input_markdown(path: str) -> tuple[str, int]:
    """Read a Markdown file, return (content, byte_size)."""
    p = Path(path)
    # ``newline=""`` recognizes every supported line ending without
    # translating it. The fenced-source occurrence carrier binds exact
    # authored EOLs before any Markdown normalization is allowed.
    with p.open("r", encoding="utf-8", newline="") as handle:
        content = handle.read()
    return content, p.stat().st_size


# ── Heading numbering removal/addition ─────────────────────────────────


def remove_md_numbering(content: str, *, rules: Any = ()) -> str:
    """Remove heading numbering patterns from Markdown headings."""
    lines = content.splitlines(keepends=True)
    result: list[str] = []
    in_fence = False
    fence_marker = ""
    for line in lines:
        line_body = line.rstrip("\r\n")
        line_ending = line[len(line_body) :]
        stripped = line_body.lstrip()
        opening_fence = _opening_fence(line_body) if not in_fence else None
        if opening_fence is not None:
            in_fence = True
            fence_marker = opening_fence
            result.append(line)
            continue
        if in_fence and _is_closing_fence(line_body, fence_marker):
            in_fence = False
            fence_marker = ""
            result.append(line)
            continue
        if in_fence:
            result.append(line)
            continue
        if line_body.startswith(("    ", "\t")):
            result.append(line)
            continue
        if stripped.startswith("#"):
            h_match = re.match(r"^(#+)\s+(.*)$", stripped)
            if h_match:
                heading_text = h_match.group(2)
                _, heading_text = strip_heading_prefix(heading_text, rules=rules)
                heading_text = heading_text.lstrip()
                indent = line_body[: len(line_body) - len(stripped)]
                result.append(f"{indent}{h_match.group(1)} {heading_text}{line_ending}")
                continue
        result.append(line)
    return "".join(result)


def add_md_numbering(
    content: str,
    scheme: str = "",
    registry: Any = None,
) -> str:
    """Add heading numbering using one exact configured scheme.

    Look up *scheme* through *registry* (a ``NumberingSchemeRegistry``),
    create a ``HeadingFormatter``, and apply it to all heading lines.

    Args:
        content: Raw Markdown content.
        scheme: Exact scheme ID (e.g. ``"gongwen_standard"``).
        registry: Request-owned numbering scheme registry.

    Returns:
        Markdown content with heading numbers inserted.
    """
    scheme_config = resolve_heading_numbering_scheme(scheme, registry)
    formatter = HeadingFormatter(scheme_config)
    return _apply_numbering(content, formatter)


def _apply_numbering(content: str, formatter: Any) -> str:
    """Apply a ``HeadingFormatter`` to all heading lines in Markdown content.

    Splits *content* by newlines, processes lines starting with ``#``,
    and returns the full text with numbering inserted.
    """
    lines = content.split("\n")
    result: list[str] = []
    in_fence = False
    fence_marker = ""
    for line in lines:
        stripped = line.lstrip()
        opening_fence = _opening_fence(line) if not in_fence else None
        if opening_fence is not None:
            in_fence = True
            fence_marker = opening_fence
            result.append(line)
            continue
        if in_fence and _is_closing_fence(line, fence_marker):
            in_fence = False
            fence_marker = ""
            result.append(line)
            continue
        if in_fence:
            result.append(line)
            continue
        if line.startswith(("    ", "\t")):
            result.append(line)
            continue
        if stripped.startswith("#"):
            # Count heading level
            level = 0
            for ch in stripped:
                if ch == "#":
                    level += 1
                else:
                    break
            heading_text = stripped[level:].strip()
            indent = line[: len(line) - len(stripped)]
            formatted = formatter.format_heading(heading_text, level)
            result.append(f"{indent}{'#' * level} {formatted}")
        else:
            result.append(line)
    return "\n".join(result)


# ── MD → Spreadsheet helpers ──────────────────────────────────────────


def parse_md_tables(content: str, *, structural_tables: bool = False) -> list[dict[str, Any]]:
    """Parse Markdown content and extract spreadsheet-facing table data.

    When structural_tables is enabled, the parser also accepts the Structural
    Tables delimiter grammar: zero or multiple header rows, an optional row
    header boundary, one-or-more delimiter dashes, and merge markers.
    """
    return parse_raw_md_tables(
        content,
        preserve_merge_marker_escapes=False,
        structural_tables=structural_tables,
    )


def parse_raw_md_tables(
    content: str,
    *,
    preserve_merge_marker_escapes: bool = False,
    structural_tables: bool = False,
) -> list[dict[str, Any]]:
    """Parse pipe tables while preserving spreadsheet-facing cell text.

    Mistune is the right parser for DOCX rendering, but spreadsheet export
    treats table cell text as Markdown source text: links remain source text,
    code spans keep their backticks, and pipes inside code spans do not split
    columns. Fenced code blocks are ignored, including unclosed fences.

    Structural Tables recognition is deliberately opt-in so the disabled
    dialect keeps its documented ordinary-Markdown behavior.
    """
    tables: list[dict[str, Any]] = []
    lines = content.splitlines()
    index = 0
    in_fence = False
    fence_marker = ""
    in_comment = False

    while index < len(lines):
        line = lines[index]
        comment_boundary = re.fullmatch(r" {0,3}%%[ \t]*", re.sub(r"^(?: {0,3}>[ \t]?)+", "", line))
        if in_comment:
            if comment_boundary is not None:
                in_comment = False
            index += 1
            continue
        opening_fence = _opening_fence(line) if not in_fence else None
        if opening_fence is not None:
            in_fence = True
            fence_marker = opening_fence
            index += 1
            continue
        if in_fence and _is_closing_fence(line, fence_marker):
            in_fence = False
            fence_marker = ""
            index += 1
            continue
        if in_fence:
            index += 1
            continue

        if comment_boundary is not None:
            in_comment = True
            index += 1
            continue

        nested_list_table = _is_nested_list_table_start(lines, index)
        if line.startswith(("    ", "\t")) and not nested_list_table:
            index += 1
            continue

        if structural_tables:
            parsed_structural = _parse_structural_table_rows(
                lines,
                index,
                nested_list_table=nested_list_table,
                preserve_merge_marker_escapes=preserve_merge_marker_escapes,
            )
            if parsed_structural is not None:
                table, index = parsed_structural
                tables.append(table)
                continue

        if index + 1 >= len(lines):
            index += 1
            continue
        header = _split_md_table_row(lines[index])
        separator = _split_md_table_row(lines[index + 1])
        table_indent = _leading_indent_columns(line) if nested_list_table else 0
        if header is None or separator is None or not _is_md_table_separator(separator):
            index += 1
            continue
        if nested_list_table and _leading_indent_columns(lines[index + 1]) < table_indent:
            index += 1
            continue

        rows: list[list[str]] = []
        index += 2
        while index < len(lines):
            if lines[index].startswith(("    ", "\t")) and not nested_list_table:
                break
            if nested_list_table and _leading_indent_columns(lines[index]) < table_indent:
                break
            if _opening_fence(lines[index]) is not None:
                break
            row = _split_md_table_row(lines[index])
            if row is None:
                break
            rows.append([_restore_markdown_table_cell(cell.strip(), preserve_merge_marker_escapes) for cell in row])
            index += 1

        if rows:
            restored_header = [
                _restore_markdown_table_cell(cell.strip(), preserve_merge_marker_escapes) for cell in header
            ]
            tables.append(
                {
                    "headers": restored_header,
                    "rows": rows,
                }
            )

    return tables


def _parse_structural_table_rows(
    lines: list[str],
    start: int,
    *,
    nested_list_table: bool,
    preserve_merge_marker_escapes: bool,
) -> tuple[dict[str, Any], int] | None:
    """Parse one structurally distinctive Structural Tables block."""

    table_indent = _leading_indent_columns(lines[start]) if nested_list_table else 0
    raw_rows: list[list[str]] = []
    cursor = start
    while cursor < len(lines):
        line = lines[cursor]
        if line.startswith(("    ", "\t")) and not nested_list_table:
            break
        if nested_list_table and _leading_indent_columns(line) < table_indent:
            break
        if _opening_fence(line) is not None:
            break
        row = _split_md_table_row(line)
        if row is None:
            break
        raw_rows.append(row)
        cursor += 1

    if len(raw_rows) < 2:
        return None

    delimiters: list[tuple[int, int, int, bool]] = []
    for row_index, row in enumerate(raw_rows):
        shape = _structural_separator_shape(row)
        if shape is None:
            continue
        column_count, header_columns, short_delimiter = shape
        delimiters.append((row_index, column_count, header_columns, short_delimiter))
    if len(delimiters) != 1:
        return None

    delimiter_index, column_count, header_columns, short_delimiter = delimiters[0]
    content_rows = [*raw_rows[:delimiter_index], *raw_rows[delimiter_index + 1 :]]
    if not content_rows or any(len(row) != column_count for row in content_rows):
        return None

    marker_found = any(cell.strip() in {"<", "^"} for row in content_rows for cell in row)
    structurally_distinctive = delimiter_index != 1 or header_columns > 0 or short_delimiter or marker_found
    if not structurally_distinctive:
        return None

    restored_rows = [
        [_restore_markdown_table_cell(cell.strip(), preserve_merge_marker_escapes) for cell in row]
        for row in content_rows
    ]
    header_rows = delimiter_index
    headers = restored_rows[header_rows - 1] if header_rows > 0 else []
    body_rows = restored_rows[header_rows:] if header_rows > 0 else restored_rows

    return (
        {
            "headers": headers,
            "rows": body_rows,
            "all_rows": restored_rows,
            "header_rows": header_rows,
            "header_columns": header_columns,
            "structural": True,
        },
        cursor,
    )


def _structural_separator_shape(cells: list[str]) -> tuple[int, int, bool] | None:
    """Return column count, row-header boundary, and short-delimiter status."""

    empty_boundaries = [index for index, cell in enumerate(cells) if cell == ""]
    if len(empty_boundaries) > 1:
        return None
    boundary = empty_boundaries[0] if empty_boundaries else None
    if boundary is not None and boundary in {0, len(cells) - 1}:
        return None

    tokens = [cell for index, cell in enumerate(cells) if index != boundary]
    if not tokens:
        return None

    dash_lengths: list[int] = []
    for token in tokens:
        stripped = token.strip()
        match = re.fullmatch(r":?(-+):?", stripped)
        if match is None:
            return None
        dash_lengths.append(len(match.group(1)))

    return len(tokens), (0 if boundary is None else boundary), any(length < 3 for length in dash_lengths)


def _leading_indent_columns(line: str) -> int:
    columns = 0
    for char in line:
        if char == " ":
            columns += 1
        elif char == "\t":
            columns += 4 - (columns % 4)
        else:
            break
    return columns


def _is_nested_list_table_start(lines: list[str], index: int) -> bool:
    """Distinguish a four-space list continuation from root indented code."""
    if index <= 0 or _leading_indent_columns(lines[index]) < 4:
        return False
    return (
        re.match(
            r"^ {0,3}(?:[-+*]|[0-9]{1,9}[.)])[ \t]+\S",
            lines[index - 1],
        )
        is not None
    )


def _opening_fence(line: str) -> str | None:
    """Return a valid fenced-code opener, preserving its exact length."""
    normalized = re.sub(r"^(?: {0,3}>[ \t]?)+", "", line)
    match = re.match(r"^ {0,3}(`{3,}|~{3,})(.*)$", normalized.rstrip("\r"))
    if match is None:
        return None
    marker = match.group(1)
    info = match.group(2)
    if marker[0] == "`" and "`" in info:
        return None
    return marker


def _is_closing_fence(line: str, opening_fence: str) -> bool:
    """Return whether *line* strictly closes *opening_fence*."""
    normalized = re.sub(r"^(?: {0,3}>[ \t]?)+", "", line)
    match = re.fullmatch(r" {0,3}(`{3,}|~{3,})[ \t]*", normalized.rstrip("\r"))
    if match is None:
        return False
    candidate = match.group(1)
    return candidate[0] == opening_fence[0] and len(candidate) >= len(opening_fence)


def _split_md_table_row(line: str) -> list[str] | None:
    stripped = re.sub(r"^(?: {0,3}>[ \t]?)+", "", line).strip()
    if "|" not in stripped:
        return None

    cells: list[str] = []
    cell_chars: list[str] = []
    escaped = False
    code_tick_count = 0
    index = 0
    while index < len(stripped):
        char = stripped[index]
        if char == "`" and (code_tick_count or not escaped):
            tick_count = _count_repeated(stripped, index, "`")
            if code_tick_count == 0:
                if _has_exact_backtick_closer(stripped, index + tick_count, tick_count):
                    code_tick_count = tick_count
            elif tick_count == code_tick_count:
                code_tick_count = 0
            cell_chars.append("`" * tick_count)
            index += tick_count
            continue
        if code_tick_count:
            cell_chars.append(char)
            index += 1
            continue
        if escaped:
            if char == "|":
                cell_chars.append(char)
            else:
                cell_chars.append("\\")
                cell_chars.append(char)
            escaped = False
            index += 1
            continue
        if char == "\\":
            escaped = True
            index += 1
            continue
        if char == "|" and code_tick_count == 0:
            cells.append("".join(cell_chars))
            cell_chars = []
            index += 1
            continue
        cell_chars.append(char)
        index += 1

    if escaped:
        cell_chars.append("\\")
    cells.append("".join(cell_chars))

    if stripped.startswith("|"):
        cells = cells[1:]
    if stripped.endswith("|"):
        cells = cells[:-1]
    return cells or None


def _has_exact_backtick_closer(text: str, start: int, length: int) -> bool:
    cursor = start
    while cursor < len(text):
        tick = text.find("`", cursor)
        if tick < 0:
            return False
        run = _count_repeated(text, tick, "`")
        if run == length:
            return True
        cursor = tick + run
    return False


def _count_repeated(text: str, start: int, char: str) -> int:
    count = 0
    while start + count < len(text) and text[start + count] == char:
        count += 1
    return count


def _is_md_table_separator(cells: list[str]) -> bool:
    if not cells:
        return False
    return all(re.fullmatch(r"\s*:?-{3,}:?\s*", cell) for cell in cells)


def _restore_markdown_table_cell(text: str, preserve_merge_marker_escapes: bool) -> str:
    restored = text.replace(r"\|", "|")
    if preserve_merge_marker_escapes:
        return _restore_table_safe_cell_text(restored)
    return _restore_table_safe_cell_text(restored.replace(r"\<", "<").replace(r"\^", "^"))


def _restore_table_safe_cell_text(text: str) -> str:
    restored = restore_table_safe_breaks(text)
    autolink_match = re.fullmatch(r"<([^<>\s]+@[^<>\s]+|https?://[^<>\s]+)>", restored)
    if autolink_match:
        return autolink_match.group(1)
    return restored


def write_table_to_csv(table_data: dict[str, Any], path: str) -> None:
    """Write a parsed table to a CSV file."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        rows = table_data.get("all_rows")
        if rows is None:
            rows = [table_data["headers"], *table_data["rows"]] if table_data["headers"] else table_data["rows"]
        writer.writerows(rows)
