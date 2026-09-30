"""Shared role and visible-projection helpers for structured clipboard tables."""

from __future__ import annotations

from docwen_core.models.clipboard_document import (
    ClipboardHardBreak,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardTable,
    ClipboardTableCell,
    ClipboardText,
    clipboard_table_header_shape,
)

_SEMANTIC_COLUMNS = (
    "Table",
    "Sheet",
    "Header Rows",
    "Header Columns",
    "Anchor",
    "Row Span",
    "Column Span",
    "Role",
    "Header",
    "Scope",
    "HTML ID",
    "Headers",
)


def _cell_role(cell: ClipboardTableCell, header_rows: int, header_columns: int) -> str:
    in_rows = header_rows > 0 and cell.row + cell.row_span <= header_rows
    in_columns = header_columns > 0 and cell.column + cell.column_span <= header_columns
    if in_rows and in_columns:
        return "corner_header"
    if in_rows:
        return "column_header"
    if in_columns:
        return "row_header"
    return "associated_header" if cell.header else "data"


def native_cell_role(table: ClipboardTable, cell: ClipboardTableCell) -> str:
    header_rows, header_columns = clipboard_table_header_shape(table)
    return _cell_role(cell, header_rows, header_columns)


def table_semantics_rows(
    tables: list[tuple[str, ClipboardTable, str, str]],
    table_titles: dict[str, str] | None = None,
) -> list[tuple[str, ...]]:
    titles = table_titles or {}
    rows: list[tuple[str, ...]] = []
    for table_id, table, _parent, _anchor in tables:
        header_rows, header_columns = clipboard_table_header_shape(table)
        sheet = titles.get(table_id, f"Table {table_id[1:]}")
        for cell in sorted(table.cells, key=lambda item: (item.row, item.column)):
            rows.append(
                (
                    table_id,
                    sheet,
                    str(header_rows),
                    str(header_columns),
                    f"R{cell.row + 1}C{cell.column + 1}",
                    str(cell.row_span),
                    str(cell.column_span),
                    _cell_role(cell, header_rows, header_columns),
                    "1" if cell.header else "0",
                    cell.scope,
                    cell.html_id,
                    " ".join(cell.headers),
                )
            )
    return rows


def semantics_headers() -> tuple[str, ...]:
    return _SEMANTIC_COLUMNS


def _simple_cell_text(cell: ClipboardTableCell) -> str | None:
    if not cell.blocks:
        return ""
    if len(cell.blocks) != 1 or not isinstance(cell.blocks[0], ClipboardParagraph):
        return None
    parts: list[str] = []
    for inline in cell.blocks[0].inlines:
        if isinstance(inline, ClipboardText):
            parts.append(inline.value)
        elif isinstance(inline, (ClipboardHardBreak, ClipboardImageRef)):
            return None
        else:
            return None
    value = "".join(parts)
    if value != value.strip():
        return None
    if "\n" in value or "\r" in value:
        return None
    return value


def _escape_structural_value(value: str) -> str:
    escaped: list[str] = []
    markdown_punctuation = frozenset("\\`*_{}[]()#+-.!|>~")
    for character in value:
        if character in markdown_punctuation or character in {"<", "^"}:
            escaped.append("\\")
        escaped.append(character)
    return "".join(escaped)


def _structural_matrix(table: ClipboardTable) -> list[list[str]] | None:
    matrix = [["" for _column in range(table.column_count)] for _row in range(table.row_count)]
    for cell in table.cells:
        value = _simple_cell_text(cell)
        if value is None:
            return None
        matrix[cell.row][cell.column] = _escape_structural_value(value)
        for row in range(cell.row, cell.row + cell.row_span):
            for column in range(cell.column, cell.column + cell.column_span):
                if row == cell.row and column == cell.column:
                    continue
                matrix[row][column] = "<" if row == cell.row else "^"
    return matrix


def structural_table_markdown(table: ClipboardTable) -> tuple[str | None, str]:
    """Return existing Structural Tables syntax only when it is lossless enough."""

    header_rows, header_columns = clipboard_table_header_shape(table)
    if header_rows <= 0:
        return None, "no_contiguous_column_header_rows"
    matrix = _structural_matrix(table)
    if matrix is None:
        return None, "recursive_or_whitespace_sensitive_cell_content"

    lines: list[str] = []
    for row_index, row in enumerate(matrix):
        if row_index == header_rows:
            delimiters = ["---"] * table.column_count
            if 0 < header_columns < table.column_count:
                lines.append(
                    "| "
                    + " | ".join(delimiters[:header_columns])
                    + " || "
                    + " | ".join(delimiters[header_columns:])
                    + " |"
                )
            else:
                lines.append("| " + " | ".join(delimiters) + " |")
        lines.append("| " + " | ".join(row) + " |")
    if header_rows == table.row_count:
        delimiters = ["---"] * table.column_count
        if 0 < header_columns < table.column_count:
            lines.append(
                "| "
                + " | ".join(delimiters[:header_columns])
                + " || "
                + " | ".join(delimiters[header_columns:])
                + " |"
            )
        else:
            lines.append("| " + " | ".join(delimiters) + " |")
    return "\n".join(lines), ""


def table_semantics_text(table_id: str, table: ClipboardTable, *, sheet: str = "") -> str:
    header_rows, header_columns = clipboard_table_header_shape(table)
    lines = [
        (
            f"table={table_id} sheet={sheet or '-'} rows={table.row_count} columns={table.column_count} "
            f"header_rows={header_rows} header_columns={header_columns}"
        )
    ]
    for cell in sorted(table.cells, key=lambda item: (item.row, item.column)):
        lines.append(
            " ".join(
                (
                    f"anchor=R{cell.row + 1}C{cell.column + 1}",
                    f"span={cell.row_span}x{cell.column_span}",
                    f"role={_cell_role(cell, header_rows, header_columns)}",
                    f"header={1 if cell.header else 0}",
                    f"scope={cell.scope or '-'}",
                    f"id={cell.html_id or '-'}",
                    f"headers={' '.join(cell.headers) or '-'}",
                )
            )
        )
    return "\n".join(lines)


def has_html_header_associations(table: ClipboardTable) -> bool:
    return any(cell.scope or cell.html_id or cell.headers for cell in table.cells)


__all__ = [
    "has_html_header_associations",
    "native_cell_role",
    "semantics_headers",
    "structural_table_markdown",
    "table_semantics_rows",
    "table_semantics_text",
]
