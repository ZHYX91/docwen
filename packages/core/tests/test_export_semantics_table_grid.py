"""Tests for ``docwen_core.export_semantics`` table semantic grid.

Covers: F-H3b-027 (TableSemanticCell, TableMergeRegion), F-H3b-029
(render_table_semantic_grid, build_table_semantic_grid).
"""

from __future__ import annotations

import pytest

from docwen_core.export_semantics import (
    TableMergeRegion,
    TableSemanticCell,
    build_table_semantic_grid,
    render_table_semantic_grid,
)

pytestmark = pytest.mark.unit

# ═══════════════════════════════════════════════════════════════════════════
# TableMergeRegion
# ═══════════════════════════════════════════════════════════════════════════


class TestTableMergeRegion:
    """F-H3b-027: ``TableMergeRegion`` frozen dataclass with rowspan/colspan."""

    def test_construction_and_properties(self) -> None:
        r = TableMergeRegion(start_row=0, start_col=1, end_row=2, end_col=3)
        assert r.start_row == 0
        assert r.start_col == 1
        assert r.end_row == 2
        assert r.end_col == 3
        assert r.rowspan == 3  # rows 0,1,2
        assert r.colspan == 3  # cols 1,2,3

    def test_single_cell_region(self) -> None:
        r = TableMergeRegion(start_row=5, start_col=5, end_row=5, end_col=5)
        assert r.rowspan == 1
        assert r.colspan == 1

    def test_frozen(self) -> None:
        r = TableMergeRegion(start_row=0, start_col=0, end_row=1, end_col=1)
        with pytest.raises(Exception):  # noqa: B017
            r.start_row = 99  # type: ignore[misc]

    def test_equality(self) -> None:
        a = TableMergeRegion(0, 0, 2, 2)
        b = TableMergeRegion(0, 0, 2, 2)
        c = TableMergeRegion(0, 0, 1, 1)
        assert a == b
        assert a != c


# ═══════════════════════════════════════════════════════════════════════════
# TableSemanticCell
# ═══════════════════════════════════════════════════════════════════════════


class TestTableSemanticCell:
    """F-H3b-027: ``TableSemanticCell`` frozen dataclass with marker property."""

    def test_non_covered_cell(self) -> None:
        cell = TableSemanticCell(
            row=0,
            col=0,
            raw_text="Hello",
            display_text="Hello",
            anchor_text="Hello",
            anchor_row=0,
            anchor_col=0,
            rowspan=1,
            colspan=1,
            is_anchor=False,
            is_covered=False,
        )
        assert cell.marker is None
        assert cell.display_text == "Hello"
        assert cell.raw_text == "Hello"

    def test_covered_same_row_anchor(self) -> None:
        """Covered cell on same row as anchor → marker '<'."""
        cell = TableSemanticCell(
            row=0,
            col=1,
            raw_text="",
            display_text="",
            anchor_text="Merged",
            anchor_row=0,
            anchor_col=0,
            rowspan=3,
            colspan=2,
            is_anchor=False,
            is_covered=True,
        )
        assert cell.marker == "<"

    def test_covered_different_row_anchor(self) -> None:
        """Covered cell on different row from anchor → marker '^'."""
        cell = TableSemanticCell(
            row=1,
            col=0,
            raw_text="",
            display_text="",
            anchor_text="Merged",
            anchor_row=0,
            anchor_col=0,
            rowspan=3,
            colspan=2,
            is_anchor=False,
            is_covered=True,
        )
        assert cell.marker == "^"

    def test_anchor_cell_has_no_marker(self) -> None:
        cell = TableSemanticCell(
            row=0,
            col=0,
            raw_text="Anchor",
            display_text="Anchor",
            anchor_text="Anchor",
            anchor_row=0,
            anchor_col=0,
            rowspan=3,
            colspan=2,
            is_anchor=True,
            is_covered=False,
        )
        assert cell.marker is None
        assert cell.is_anchor is True
        assert cell.is_covered is False

    def test_all_fields_present(self) -> None:
        """Verify the dataclass has all 11 fields expected from old code."""
        cell = TableSemanticCell(
            row=1,
            col=2,
            raw_text="raw",
            display_text="disp",
            anchor_text="anch",
            anchor_row=0,
            anchor_col=0,
            rowspan=2,
            colspan=3,
            is_anchor=False,
            is_covered=True,
        )
        assert cell.row == 1
        assert cell.col == 2
        assert cell.raw_text == "raw"
        assert cell.display_text == "disp"
        assert cell.anchor_text == "anch"
        assert cell.anchor_row == 0
        assert cell.anchor_col == 0
        assert cell.rowspan == 2
        assert cell.colspan == 3
        assert cell.is_anchor is False
        assert cell.is_covered is True

    def test_frozen(self) -> None:
        cell = TableSemanticCell(
            row=0,
            col=0,
            raw_text="x",
            display_text="x",
            anchor_text="x",
            anchor_row=0,
            anchor_col=0,
            rowspan=1,
            colspan=1,
            is_anchor=False,
            is_covered=False,
        )
        with pytest.raises(Exception):  # noqa: B017
            cell.raw_text = "y"  # type: ignore[misc]


# ═══════════════════════════════════════════════════════════════════════════
# build_table_semantic_grid
# ═══════════════════════════════════════════════════════════════════════════


class TestBuildTableSemanticGrid:
    """F-H3b-027: ``build_table_semantic_grid`` constructs a 2‑D semantic grid."""

    def test_no_merges_flat_grid(self) -> None:
        """Without merge regions, every cell is independent."""
        cell_text = {(0, 0): "A", (0, 1): "B", (1, 0): "C", (1, 1): "D"}
        grid = build_table_semantic_grid(
            row_count=2,
            col_count=2,
            cell_text_by_position=cell_text,
            merge_regions=[],
        )
        assert len(grid) == 2
        assert len(grid[0]) == 2
        for row in grid:
            for cell in row:
                assert not cell.is_covered
                assert not cell.is_anchor
                assert cell.rowspan == 1
                assert cell.colspan == 1
                assert cell.display_text == cell.raw_text

    def test_empty_cells_default_to_empty_string(self) -> None:
        """Cells not in the text map get raw_text=''."""
        grid = build_table_semantic_grid(
            row_count=1,
            col_count=2,
            cell_text_by_position={(0, 0): "X"},
            merge_regions=[],
        )
        assert grid[0][0].raw_text == "X"
        assert grid[0][1].raw_text == ""

    def test_single_merge_region_2x2(self) -> None:
        """A 2×2 merge region produces correct anchor / covered metadata."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=1, end_col=1)]
        cell_text = {(0, 0): "Merged"}
        grid = build_table_semantic_grid(
            row_count=2,
            col_count=2,
            cell_text_by_position=cell_text,
            merge_regions=merge,
        )
        # Anchor cell
        anchor = grid[0][0]
        assert anchor.is_anchor is True
        assert anchor.is_covered is False
        assert anchor.anchor_text == "Merged"
        assert anchor.display_text == "Merged"
        assert anchor.rowspan == 2
        assert anchor.colspan == 2
        assert anchor.marker is None

        # Covered cells
        for rc in [(0, 1), (1, 0), (1, 1)]:
            cell = grid[rc[0]][rc[1]]
            assert cell.is_covered is True
            assert cell.is_anchor is False
            assert cell.anchor_text == "Merged"
            assert cell.anchor_row == 0
            assert cell.anchor_col == 0

    def test_multiple_disjoint_merge_regions(self) -> None:
        """Two non-overlapping merge regions."""
        regions = [
            TableMergeRegion(start_row=0, start_col=0, end_row=0, end_col=1),
            TableMergeRegion(start_row=1, start_col=0, end_row=1, end_col=1),
        ]
        cell_text = {(0, 0): "Top", (1, 0): "Bottom"}
        grid = build_table_semantic_grid(
            row_count=2,
            col_count=2,
            cell_text_by_position=cell_text,
            merge_regions=regions,
        )
        assert grid[0][0].is_anchor is True
        assert grid[0][0].anchor_text == "Top"
        assert grid[0][1].is_covered is True
        assert grid[0][1].anchor_text == "Top"

        assert grid[1][0].is_anchor is True
        assert grid[1][0].anchor_text == "Bottom"
        assert grid[1][1].is_covered is True
        assert grid[1][1].anchor_text == "Bottom"

    def test_merge_with_empty_anchor_text(self) -> None:
        """When the anchor position has no text, anchor_text is ''."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=1, end_col=1)]
        grid = build_table_semantic_grid(
            row_count=2,
            col_count=2,
            cell_text_by_position={},
            merge_regions=merge,
        )
        assert grid[0][0].anchor_text == ""
        assert grid[0][1].anchor_text == ""

    def test_row_count_and_col_count_respected(self) -> None:
        grid = build_table_semantic_grid(
            row_count=3,
            col_count=5,
            cell_text_by_position={},
            merge_regions=[],
        )
        assert len(grid) == 3
        assert all(len(row) == 5 for row in grid)

    def test_marker_property_on_grid_cells(self) -> None:
        """marker property on built grid cells matches specification."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=2, end_col=1)]
        cell_text = {(0, 0): "A"}
        grid = build_table_semantic_grid(
            row_count=3,
            col_count=2,
            cell_text_by_position=cell_text,
            merge_regions=merge,
        )
        # Same row as anchor → "<"
        assert grid[0][1].marker == "<"
        # Different row from anchor → "^"
        assert grid[1][0].marker == "^"
        assert grid[1][1].marker == "^"
        assert grid[2][0].marker == "^"
        # Anchor has no marker
        assert grid[0][0].marker is None


# ═══════════════════════════════════════════════════════════════════════════
# render_table_semantic_grid
# ═══════════════════════════════════════════════════════════════════════════


def _make_grid(
    row_count: int,
    col_count: int,
    cell_text: dict[tuple[int, int], str],
    merge_regions: list[TableMergeRegion],
) -> list[list[TableSemanticCell]]:
    return build_table_semantic_grid(
        row_count=row_count,
        col_count=col_count,
        cell_text_by_position=cell_text,
        merge_regions=merge_regions,
    )


class TestRenderTableSemanticGrid:
    """F-H3b-029: ``render_table_semantic_grid`` applies strategy to grid."""

    # ── fill strategy ────────────────────────────────────────────────

    def test_fill_strategy_expands_anchor_text(self) -> None:
        """With 'fill' strategy, covered cells get the anchor text."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=1, end_col=1)]
        grid = _make_grid(2, 2, {(0, 0): "Hello"}, merge)
        result = render_table_semantic_grid(grid, strategy="fill")
        assert result == [["Hello", "Hello"], ["Hello", "Hello"]]

    def test_fill_strategy_non_merged_cells_unchanged(self) -> None:
        """Non-merged cells keep their display_text."""
        grid = _make_grid(2, 2, {(0, 0): "A", (0, 1): "B", (1, 0): "C", (1, 1): "D"}, [])
        result = render_table_semantic_grid(grid, strategy="fill")
        assert result == [["A", "B"], ["C", "D"]]

    def test_fill_deals_with_empty_anchor(self) -> None:
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=0, end_col=1)]
        grid = _make_grid(1, 2, {}, merge)
        result = render_table_semantic_grid(grid, strategy="fill")
        assert result == [["", ""]]

    def test_removed_replicate_alias_uses_generic_invalid_default(self) -> None:
        """The removed alias follows the same default path as any invalid value."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=1, end_col=1)]
        grid = _make_grid(2, 2, {(0, 0): "X"}, merge)
        result_fill = render_table_semantic_grid(grid, strategy="fill")
        result_invalid = render_table_semantic_grid(grid, strategy="replicate")
        assert result_invalid == result_fill == [["X", "X"], ["X", "X"]]

    # ── marker strategy ──────────────────────────────────────────────

    def test_marker_strategy_uses_markers(self) -> None:
        """Covered cells get '<' (same row) or '^' (different row)."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=2, end_col=1)]
        grid = _make_grid(3, 2, {(0, 0): "Anchor"}, merge)
        result = render_table_semantic_grid(grid, strategy="marker")
        # Row 0: anchor col gets "Anchor"; covered col 1 gets "<" (same row)
        # Row 1: both cols covered, get "^" (different row)
        # Row 2: both cols covered, get "^"
        assert result == [
            ["Anchor", "<"],
            ["^", "^"],
            ["^", "^"],
        ]

    def test_marker_strategy_non_merged_unchanged(self) -> None:
        grid = _make_grid(1, 3, {(0, 0): "A", (0, 1): "B", (0, 2): "C"}, [])
        result = render_table_semantic_grid(grid, strategy="marker")
        assert result == [["A", "B", "C"]]

    def test_marker_strategy_mixed_merged_and_non_merged(self) -> None:
        """Merge region covers only part of the row; the rest is independent."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=0, end_col=1)]
        grid = _make_grid(1, 3, {(0, 0): "M", (0, 2): "Free"}, merge)
        result = render_table_semantic_grid(grid, strategy="marker")
        assert result == [["M", "<", "Free"]]

    # ── empty strategy ───────────────────────────────────────────────

    def test_empty_strategy_covered_become_empty(self) -> None:
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=1, end_col=1)]
        grid = _make_grid(2, 2, {(0, 0): "Value"}, merge)
        result = render_table_semantic_grid(grid, strategy="empty")
        assert result == [["Value", ""], ["", ""]]

    def test_empty_strategy_non_merged_unchanged(self) -> None:
        grid = _make_grid(2, 2, {(0, 0): "A", (0, 1): "B"}, [])
        result = render_table_semantic_grid(grid, strategy="empty")
        assert result == [["A", "B"], ["", ""]]

    # ── round-trip sanity ────────────────────────────────────────────

    def test_strategies_produce_expected_dimensions(self) -> None:
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=2, end_col=2)]
        grid = _make_grid(3, 3, {(0, 0): "Big"}, merge)
        for strategy in ("fill", "marker", "empty"):
            result = render_table_semantic_grid(grid, strategy=strategy)
            assert len(result) == 3
            assert all(len(row) == 3 for row in result)

    def test_large_grid_no_merges(self) -> None:
        """A 10×10 grid with no merges should pass through unchanged."""
        cell_text = {(r, c): f"R{r}C{c}" for r in range(10) for c in range(10)}
        grid = _make_grid(10, 10, cell_text, [])
        result = render_table_semantic_grid(grid, strategy="fill")
        for r in range(10):
            for c in range(10):
                assert result[r][c] == f"R{r}C{c}"


# ═══════════════════════════════════════════════════════════════════════════
# Integration: build + render pipeline
# ═══════════════════════════════════════════════════════════════════════════


class TestBuildAndRenderPipeline:
    """End-to-end tests of build_table_semantic_grid → render_table_semantic_grid."""

    def test_realistic_spreadsheet_scenario(self) -> None:
        """Simulate a spreadsheet with one merge region and standalone cells."""
        # A 4×4 grid:
        #   [A   ][B]
        #   [     ][C]
        #   [D][E ][F]
        #   [G][  ][I]
        # Merge (0,0)-(1,0): vertical merge of "A"
        # Merge (2,1)-(3,1): vertical merge of "E"
        merge_regions = [
            TableMergeRegion(start_row=0, start_col=0, end_row=1, end_col=0),
            TableMergeRegion(start_row=2, start_col=1, end_row=3, end_col=1),
        ]
        cell_text = {
            (0, 0): "A",
            (0, 1): "B",
            (1, 1): "C",
            (2, 0): "D",
            (2, 1): "E",
            (2, 2): "F",
            (3, 0): "G",
            (3, 2): "I",
        }
        grid = build_table_semantic_grid(
            row_count=4,
            col_count=3,
            cell_text_by_position=cell_text,
            merge_regions=merge_regions,
        )

        # --- fill strategy ---
        filled = render_table_semantic_grid(grid, strategy="fill")
        assert filled[0] == ["A", "B", ""]  # row 0: B is standalone, col 2 empty
        assert filled[1] == ["A", "C", ""]  # (1,0) covered by A-merge → fill "A"
        assert filled[2] == ["D", "E", "F"]  # anchors D, E; F standalone
        assert filled[3] == ["G", "E", "I"]  # (3,1) covered → fill "E"

        # --- marker strategy ---
        marked = render_table_semantic_grid(grid, strategy="marker")
        assert marked[0] == ["A", "B", ""]
        assert marked[1] == ["^", "C", ""]  # (1,0) different row from anchor → "^"
        assert marked[2] == ["D", "E", "F"]
        assert marked[3] == ["G", "^", "I"]  # (3,1) different row from anchor → "^"

        # --- empty strategy ---
        emptied = render_table_semantic_grid(grid, strategy="empty")
        assert emptied[0] == ["A", "B", ""]
        assert emptied[1] == ["", "C", ""]  # covered → ""
        assert emptied[2] == ["D", "E", "F"]
        assert emptied[3] == ["G", "", "I"]  # covered → ""

    def test_fully_merged_sheet(self) -> None:
        """A single merge region covering the entire sheet."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=2, end_col=2)]
        cell_text = {(0, 0): "All"}
        grid = build_table_semantic_grid(
            row_count=3,
            col_count=3,
            cell_text_by_position=cell_text,
            merge_regions=merge,
        )
        filled = render_table_semantic_grid(grid, strategy="fill")
        for r in range(3):
            for c in range(3):
                assert filled[r][c] == "All"

    def test_no_cell_text_all_empty(self) -> None:
        """When no cell has text, the grid is all empty strings."""
        merge = [TableMergeRegion(start_row=0, start_col=0, end_row=1, end_col=1)]
        grid = build_table_semantic_grid(
            row_count=2,
            col_count=2,
            cell_text_by_position={},
            merge_regions=merge,
        )
        filled = render_table_semantic_grid(grid, strategy="fill")
        assert filled == [["", ""], ["", ""]]
