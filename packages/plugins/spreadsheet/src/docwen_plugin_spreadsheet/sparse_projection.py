"""Plan bounded native spreadsheet projections from real stored coordinates."""

from __future__ import annotations

from collections import deque
from collections.abc import Callable, Mapping
from typing import Any

from docwen_core.export_semantics import TableMergeRegion, build_table_semantic_grid, render_table_semantic_grid
from docwen_core.spreadsheet_projection import ProjectionBudget, SpreadsheetProjectionLimitError
from docwen_plugin_spreadsheet.native_links import native_cell_markdown


def encode_generated_cell_fragment(markdown: str) -> str:
    """Protect table delimiters in already-generated inline Markdown.

    Authored native text must use encode_markdown_table_cell instead. Existing
    inline escapes remain unchanged, so generated links are not encoded twice.
    """
    markdown = markdown.replace("\r\n", "\n").replace("\r", "\n")
    result: list[str] = []
    backslashes = 0
    for char in markdown:
        if char == "\n":
            result.append("<br>")
        elif char == "|" and backslashes % 2 == 0:
            result.append("\\|")
        else:
            result.append(char)
        backslashes = backslashes + 1 if char == "\\" else 0
    return "".join(result)


def worksheet_blocks(
    worksheet: Any,
    *,
    strategy: str,
    generated: Mapping[tuple[int, int], list[str]],
    budget: ProjectionBudget,
    copy_budget: ProjectionBudget,
    cancel_check: Callable[[], None],
    split_rendered: Callable[..., list[Any]],
) -> list[Any]:
    """Build grids only for connected actual cell/merge/image components."""
    import pandas as pd

    raw: dict[tuple[int, int], str] = {}
    for index, cell in enumerate(worksheet._cells.values()):
        if index % 1000 == 0:
            cancel_check()
        value = cell.value
        if value is not None and (not isinstance(value, str) or value.strip()):
            raw[(cell.row - 1, cell.column - 1)] = native_cell_markdown(cell)
    positions = set(raw)
    positions.update(position for position, parts in generated.items() if any(parts))
    merges: list[TableMergeRegion] = []
    merge_budget = ProjectionBudget(limit=budget.limit)
    for native in worksheet.merged_cells.ranges:
        cancel_check()
        region = TableMergeRegion(native.min_row - 1, native.min_col - 1, native.max_row - 1, native.max_col - 1)
        merge_budget.reserve(region.rowspan * region.colspan, location=worksheet.title)
        merges.append(region)
        for row in range(region.start_row, region.end_row + 1):
            for column in range(region.start_col, region.end_col + 1):
                if ((row - region.start_row) * region.colspan + column - region.start_col) % 1000 == 0:
                    cancel_check()
                positions.add((row, column))
    # Actual positions, including merge coverage, are bounded before traversal.
    if len(positions) > budget.limit:
        raise SpreadsheetProjectionLimitError("Worksheet active coordinates exceed the projection budget.")
    unseen = positions.copy()
    components: list[set[tuple[int, int]]] = []
    for start in sorted(positions):
        if start not in unseen:
            continue
        unseen.remove(start)
        component = {start}
        queue = deque([start])
        visits = 0
        while queue:
            if visits % 1000 == 0:
                cancel_check()
            visits += 1
            row, column = queue.popleft()
            for adjacent in ((row - 1, column), (row + 1, column), (row, column - 1), (row, column + 1)):
                if adjacent in unseen:
                    unseen.remove(adjacent)
                    component.add(adjacent)
                    queue.append(adjacent)
        components.append(component)
    blocks: list[Any] = []
    merge_by_anchor = {(region.start_row, region.start_col): region for region in merges}
    for component in components:
        cancel_check()
        top = min(row for row, _ in component)
        bottom = max(row for row, _ in component)
        left = min(column for _, column in component)
        right = max(column for _, column in component)
        rows, columns = bottom - top + 1, right - left + 1
        budget.reserve(rows * columns, location=worksheet.title)
        local_raw = {
            (row - top, column - left): raw[(row, column)] for row, column in component if (row, column) in raw
        }
        local_merges = [
            TableMergeRegion(
                region.start_row - top, region.start_col - left, region.end_row - top, region.end_col - left
            )
            for position in component
            if (region := merge_by_anchor.get(position)) is not None
        ]
        grid = build_table_semantic_grid(
            row_count=rows,
            col_count=columns,
            cell_text_by_position=local_raw,
            merge_regions=local_merges,
            max_cells=budget.limit,
            cancel_check=cancel_check,
        )
        rendered = render_table_semantic_grid(grid, strategy=strategy, cancel_check=cancel_check)
        for row, column in component:
            parts = generated.get((row, column))
            if not parts:
                continue
            addition = " ".join(encode_generated_cell_fragment(part) for part in parts if part)
            current = rendered[row - top][column - left]
            rendered[row - top][column - left] = current + ("<br>" if current and addition else "") + addition
        frame = pd.DataFrame(rendered, index=range(top, bottom + 1), columns=range(left, right + 1))
        blocks.extend(split_rendered(frame, cancel_check=cancel_check, copy_budget=copy_budget))
    blocks.sort(key=lambda block: (int(block.index[0]), int(block.columns[0])))
    return blocks
