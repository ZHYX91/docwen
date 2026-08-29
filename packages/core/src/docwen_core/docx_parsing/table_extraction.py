"""Business-neutral OOXML table geometry and Markdown projection."""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_core.export_semantics import TableSemanticCell


DocxCellTextResolver = Callable[[Any, int, int], str]


class DocxTableGeometryError(ValueError):
    """One deterministic malformed-geometry failure from a DOCX table."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def build_docx_table_semantic_grid(
    tbl_element: Any,
    *,
    cell_text_resolver: DocxCellTextResolver,
) -> list[list[TableSemanticCell]]:
    """Build one merge-aware semantic grid from an OOXML ``w:tbl``.

    Consumers own cell-content rendering through ``cell_text_resolver``;
    Core owns only row/column positioning and ``gridSpan``/``vMerge`` geometry.
    """

    from docx.oxml.ns import qn

    from docwen_core.export_semantics import TableMergeRegion, build_table_semantic_grid

    rows = _logical_children(tbl_element, target_tag=qn("w:tr"))
    if not rows:
        raise DocxTableGeometryError("rows_missing", "DOCX table has no rows.")

    table_grid = tbl_element.find(qn("w:tblGrid"))
    declared_grid_width = len(table_grid.findall(qn("w:gridCol"))) if table_grid is not None else 0
    row_layouts: list[tuple[Any, list[Any], list[int], int, int]] = []
    for row_index, row in enumerate(rows):
        cells = _logical_children(row, target_tag=qn("w:tc"))
        if not cells:
            raise DocxTableGeometryError(
                "cells_missing",
                f"DOCX table row {row_index} has no cells.",
            )
        spans = [_grid_span(cell.find(qn("w:tcPr"))) for cell in cells]
        grid_before, grid_after = _row_grid_offsets(row)
        row_layouts.append((row, cells, spans, grid_before, grid_after))

    if declared_grid_width > 0:
        grid_width = max(declared_grid_width, *(sum(spans) for _row, _cells, spans, _before, _after in row_layouts))
        while True:
            effective_grid_befores = [
                grid_before if grid_before <= grid_width else 0
                for _row, _cells, _spans, grid_before, _grid_after in row_layouts
            ]
            expanded_grid_width = max(
                grid_width,
                *(
                    grid_before + sum(spans)
                    for grid_before, (_row, _cells, spans, _raw_before, _after) in zip(
                        effective_grid_befores, row_layouts, strict=True
                    )
                ),
            )
            if expanded_grid_width == grid_width:
                break
            grid_width = expanded_grid_width
    else:
        effective_grid_befores = [grid_before for _row, _cells, _spans, grid_before, _grid_after in row_layouts]
        grid_width = max(
            grid_before + sum(spans) + grid_after
            for grid_before, (_row, _cells, spans, _raw_before, grid_after) in zip(
                effective_grid_befores, row_layouts, strict=True
            )
        )

    cell_text_by_position: dict[tuple[int, int], str] = {}
    merge_regions: list[TableMergeRegion] = []
    active_vertical_merges: dict[int, tuple[int, int]] = {}

    for row_index, ((_row, cells, spans, _raw_grid_before, _grid_after), grid_before) in enumerate(
        zip(row_layouts, effective_grid_befores, strict=True)
    ):
        virtual_col = grid_before
        continued_vertical_merges: set[int] = set()
        next_vertical_merges: dict[int, tuple[int, int]] = {}
        for cell, colspan in zip(cells, spans, strict=True):
            tc_pr = cell.find(qn("w:tcPr"))
            vertical_merge = _vertical_merge_state(tc_pr)

            if vertical_merge == "continue":
                active_merge = active_vertical_merges.get(virtual_col)
                if active_merge is None or active_merge[1] != colspan:
                    raise DocxTableGeometryError(
                        "vmerge_invalid",
                        f"DOCX table row {row_index} has an orphan or mismatched vertical merge continuation.",
                    )
                continued_vertical_merges.add(virtual_col)
                next_vertical_merges[virtual_col] = active_merge
                for column in range(virtual_col, virtual_col + colspan):
                    cell_text_by_position[(row_index, column)] = ""
            else:
                cell_text = cell_text_resolver(cell, row_index, virtual_col)
                for column in range(virtual_col, virtual_col + colspan):
                    cell_text_by_position[(row_index, column)] = cell_text if column == virtual_col else ""

            if vertical_merge == "restart":
                next_vertical_merges[virtual_col] = (row_index, colspan)
            elif colspan > 1 and vertical_merge is None:
                merge_regions.append(
                    TableMergeRegion(
                        start_row=row_index,
                        start_col=virtual_col,
                        end_row=row_index,
                        end_col=virtual_col + colspan - 1,
                    )
                )

            virtual_col += colspan
        for start_col, (start_row, colspan) in active_vertical_merges.items():
            if start_col not in continued_vertical_merges:
                merge_regions.append(
                    TableMergeRegion(
                        start_row=start_row,
                        start_col=start_col,
                        end_row=row_index - 1,
                        end_col=start_col + colspan - 1,
                    )
                )
        active_vertical_merges = next_vertical_merges

    for start_col, (start_row, colspan) in active_vertical_merges.items():
        merge_regions.append(
            TableMergeRegion(
                start_row=start_row,
                start_col=start_col,
                end_row=len(rows) - 1,
                end_col=start_col + colspan - 1,
            )
        )

    if not cell_text_by_position:
        raise DocxTableGeometryError("cells_missing", "DOCX table has no usable cells.")
    return build_table_semantic_grid(
        row_count=len(rows),
        col_count=grid_width,
        cell_text_by_position=cell_text_by_position,
        merge_regions=merge_regions,
    )


def render_docx_table_rows(
    tbl_element: Any,
    *,
    cell_text_resolver: DocxCellTextResolver,
    strategy: str = "fill",
    escape_literal_merge_markers: bool = False,
) -> list[list[str]]:
    """Render an OOXML table through the shared semantic-grid policy."""

    from docwen_core.export_semantics import render_table_semantic_grid

    grid = build_docx_table_semantic_grid(
        tbl_element,
        cell_text_resolver=cell_text_resolver,
    )
    if not grid:
        return []
    rendered = render_table_semantic_grid(grid, strategy=strategy)
    if not escape_literal_merge_markers:
        return rendered
    return [
        [
            (f"\\{value}" if value in {"<", "^"} and (strategy != "marker" or not cell.is_covered) else value)
            for cell, value in zip(row_cells, rendered_row, strict=True)
        ]
        for row_cells, rendered_row in zip(grid, rendered, strict=True)
    ]


def markdown_table_lines(
    rendered: list[list[str]],
    *,
    header_rows: int = 1,
    header_columns: int = 0,
) -> list[str]:
    """Project rendered semantic rows into a Markdown table."""

    if not rendered or not any(any(cell for cell in row) for row in rendered):
        return []
    width = len(rendered[0])
    bounded_header_rows = max(1, min(header_rows, len(rendered)))
    bounded_header_columns = max(0, min(header_columns, width))
    lines: list[str] = []
    for row_index, row in enumerate(rendered):
        if row_index == bounded_header_rows:
            delimiters = ["---"] * width
            if 0 < bounded_header_columns < width:
                line = (
                    "| "
                    + " | ".join(delimiters[:bounded_header_columns])
                    + " || "
                    + " | ".join(delimiters[bounded_header_columns:])
                    + " |"
                )
            else:
                line = "| " + " | ".join(delimiters) + " |"
            lines.append(line)
        padded = [*row, *([""] * (width - len(row)))]
        lines.append("| " + " | ".join(padded[:width]) + " |")
    if bounded_header_rows == len(rendered):
        delimiters = ["---"] * width
        if 0 < bounded_header_columns < width:
            lines.append(
                "| "
                + " | ".join(delimiters[:bounded_header_columns])
                + " || "
                + " | ".join(delimiters[bounded_header_columns:])
                + " |"
            )
        else:
            lines.append("| " + " | ".join(delimiters) + " |")
    return lines


def _logical_children(parent: Any, *, target_tag: str) -> list[Any]:
    """Return ordered logical rows or cells through only legal OOXML wrappers."""

    from docx.oxml.ns import qn

    custom_xml_tag = qn("w:customXml")
    sdt_tag = qn("w:sdt")
    sdt_content_tag = qn("w:sdtContent")
    pending = list(reversed(list(parent)))
    logical_children: list[Any] = []
    while pending:
        child = pending.pop()
        if child.tag == target_tag:
            logical_children.append(child)
        elif child.tag == custom_xml_tag:
            pending.extend(reversed(list(child)))
        elif child.tag == sdt_tag:
            contents = [item for item in child if item.tag == sdt_content_tag]
            for content in reversed(contents):
                pending.extend(reversed(list(content)))
    return logical_children


def _row_grid_offsets(row: Any) -> tuple[int, int]:
    """Return validated leading/trailing row grid offsets."""

    from docx.oxml.ns import qn

    tr_pr = row.find(qn("w:trPr"))
    if tr_pr is None:
        return 0, 0
    grid_before_element = tr_pr.find(qn("w:gridBefore"))
    grid_after_element = tr_pr.find(qn("w:gridAfter"))
    grid_before = _grid_offset(grid_before_element, name="gridBefore")
    grid_after = _grid_offset(grid_after_element, name="gridAfter")
    return grid_before, grid_after


def _grid_offset(element: Any, *, name: str) -> int:
    """Return one non-negative explicit OOXML row grid offset."""

    from docx.oxml.ns import qn

    if element is None:
        return 0
    raw_value = element.get(qn("w:val"))
    try:
        offset = int(raw_value)
    except (TypeError, ValueError) as exc:
        raise DocxTableGeometryError(
            "grid_offset_invalid",
            f"DOCX table {name} must be a non-negative integer.",
        ) from exc
    if offset < 0:
        raise DocxTableGeometryError(
            "grid_offset_invalid",
            f"DOCX table {name} must be a non-negative integer.",
        )
    return offset


def _vertical_merge_state(tc_pr: Any) -> str | None:
    """Return one strict OOXML vertical-merge state."""

    from docx.oxml.ns import qn

    if tc_pr is None:
        return None
    vertical_merge = tc_pr.find(qn("w:vMerge"))
    if vertical_merge is None:
        return None
    raw_value = vertical_merge.get(qn("w:val"))
    if raw_value is None:
        return "continue"
    if raw_value not in {"continue", "restart"}:
        raise DocxTableGeometryError(
            "vmerge_invalid",
            "DOCX table vMerge must be continue, restart, or omitted.",
        )
    return raw_value


def _grid_span(tc_pr: Any) -> int:
    """Return one positive OOXML grid span or a stable geometry error."""

    from docx.oxml.ns import qn

    if tc_pr is None:
        return 1
    grid_span = tc_pr.find(qn("w:gridSpan"))
    if grid_span is None:
        return 1
    raw_value = grid_span.get(qn("w:val"))
    try:
        span = int(raw_value)
    except (TypeError, ValueError) as exc:
        raise DocxTableGeometryError(
            "grid_span_invalid",
            "DOCX table gridSpan must be a positive integer.",
        ) from exc
    if span <= 0:
        raise DocxTableGeometryError(
            "grid_span_invalid",
            "DOCX table gridSpan must be a positive integer.",
        )
    return span
