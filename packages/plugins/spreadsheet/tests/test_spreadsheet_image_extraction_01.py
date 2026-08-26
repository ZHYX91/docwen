"""Focused tests split from test_spreadsheet_image_extraction.py."""

from __future__ import annotations

from ._spreadsheet_image_extraction_support import (
    _make_wb_with_merge,
    extract_yaml,
    format_sanitized_image_link,
    generate_basic_yaml_frontmatter,
    openpyxl,
    pytest,
)


@pytest.mark.unit
class TestWorksheetToDataframeGrid:
    """Verify that ``_worksheet_to_dataframe`` produces correct DataFrames
    through the shared semantic grid pipeline."""

    def test_no_merges_simple_grid(self) -> None:
        """Without merged cells, each cell maps 1:1."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [
                ["A", "B"],
                ["C", "D"],
            ]
        )
        df = _worksheet_to_dataframe(wb.active, table_merge_strategy="fill")
        assert df.shape == (2, 2)
        assert df.iat[0, 0] == "A"
        assert df.iat[0, 1] == "B"
        assert df.iat[1, 0] == "C"
        assert df.iat[1, 1] == "D"

    def test_fill_strategy_expands_merge(self) -> None:
        """'fill' strategy fills covered cells with the anchor value."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [["Header", None], [None, None]],
            merge_ranges=["A1:B2"],
        )
        df = _worksheet_to_dataframe(wb.active, table_merge_strategy="fill")
        assert df.shape == (2, 2)
        for r in range(2):
            for c in range(2):
                assert df.iat[r, c] == "Header"

    def test_empty_strategy_leaves_covered_blank(self) -> None:
        """'empty' strategy leaves covered cells empty."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [["Only", None], [None, "Standalone"]],
            merge_ranges=["A1:A2"],
        )
        df = _worksheet_to_dataframe(wb.active, table_merge_strategy="empty")
        assert df.iat[0, 0] == "Only"
        assert df.iat[1, 0] == ""
        assert df.iat[1, 1] == "Standalone"

    def test_marker_strategy_grid(self) -> None:
        """'marker' strategy uses '<' and '^' for covered cells."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [["Anchor", None], [None, None], ["Below", "X"]],
            merge_ranges=["A1:B2"],
        )
        df = _worksheet_to_dataframe(wb.active, table_merge_strategy="marker")
        assert df.shape == (3, 2)
        assert df.iat[0, 0] == "Anchor"
        assert df.iat[0, 1] == "<"
        assert df.iat[1, 0] == "^"
        assert df.iat[1, 1] == "^"
        assert df.iat[2, 0] == "Below"
        assert df.iat[2, 1] == "X"

    def test_removed_replicate_strategy_uses_default(self) -> None:
        """The removed alias is treated as an invalid value, using the default."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [["X", None], [None, "Y"]],
            merge_ranges=["A1:A2"],
        )
        df_fill = _worksheet_to_dataframe(wb.active, table_merge_strategy="fill")
        wb2 = _make_wb_with_merge(
            [["X", None], [None, "Y"]],
            merge_ranges=["A1:A2"],
        )
        df_invalid = _worksheet_to_dataframe(wb2.active, table_merge_strategy="replicate")
        assert df_fill.equals(df_invalid)
        assert df_fill.iat[0, 0] == "X"
        assert df_fill.iat[1, 0] == "X"

    def test_multiple_disjoint_merges_grid(self) -> None:
        """Two non-overlapping merge regions in the same sheet."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [
                ["Top", None, "Right"],
                [None, "Mid", None],
            ],
            merge_ranges=["A1:A2", "C1:C2"],
        )
        df = _worksheet_to_dataframe(wb.active, table_merge_strategy="fill")
        assert df.iat[0, 0] == "Top"
        assert df.iat[1, 0] == "Top"
        assert df.iat[0, 1] == ""
        assert df.iat[1, 1] == "Mid"
        assert df.iat[0, 2] == "Right"
        assert df.iat[1, 2] == "Right"

    def test_empty_sheet_returns_empty_dataframe(self) -> None:
        """An empty worksheet produces an empty DataFrame."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = openpyxl.Workbook()
        df = _worksheet_to_dataframe(wb.active)
        assert df.empty

    def test_sheet_with_only_empty_cells(self) -> None:
        """All cells are None/empty -> empty DataFrame."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [
                [None, None],
                [None, None],
            ]
        )
        df = _worksheet_to_dataframe(wb.active)
        assert df.empty

    def test_default_strategy_is_fill(self) -> None:
        """When no strategy is given, 'fill' is the default."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [["Val", None], [None, None]],
            merge_ranges=["A1:B2"],
        )
        df = _worksheet_to_dataframe(wb.active)
        for r in range(2):
            for c in range(2):
                assert df.iat[r, c] == "Val"

    def test_invalid_strategy_falls_back_to_fill(self) -> None:
        """An unrecognized strategy name normalizes to 'fill'."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [["X", None], [None, "Y"]],
            merge_ranges=["A1:B1"],
        )
        df = _worksheet_to_dataframe(wb.active, table_merge_strategy="garbage")
        assert df.iat[0, 0] == "X"
        assert df.iat[0, 1] == "X"
        assert df.iat[1, 0] == ""
        assert df.iat[1, 1] == "Y"

    def test_numeric_cell_values_preserved(self) -> None:
        """Numeric cell values should survive the grid pipeline."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = _make_wb_with_merge(
            [
                [42, None],
                [None, 3.14],
            ]
        )
        df = _worksheet_to_dataframe(wb.active)
        assert df.iat[0, 0] == "42"
        assert df.iat[1, 1] == "3.14"

    def test_merge_with_empty_anchor(self) -> None:
        """Merge region whose anchor cell has no value."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            _worksheet_to_dataframe,
        )

        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws.merge_cells("A1:B2")
        ws.cell(row=1, column=3, value="Only standalone")
        df = _worksheet_to_dataframe(ws, table_merge_strategy="fill")
        assert df.iat[0, 0] == ""
        assert df.iat[0, 1] == ""
        assert df.iat[1, 0] == ""
        assert df.iat[1, 1] == ""
        assert df.iat[0, 2] == "Only standalone"


@pytest.mark.unit
class TestSpreadsheetYamlFrontmatter:
    """Verify that the spreadsheet converter produces correct YAML frontmatter
    via the shared ``generate_basic_yaml_frontmatter`` utility."""

    def test_yaml_contains_title_and_aliases(self) -> None:
        frontmatter = generate_basic_yaml_frontmatter("季度报表")
        assert "title: 季度报表" in frontmatter
        assert "aliases:" in frontmatter
        assert "  - 季度报表" in frontmatter

    def test_yaml_starts_and_ends_correctly(self) -> None:
        frontmatter = generate_basic_yaml_frontmatter("TestSheet")
        assert frontmatter.startswith("---\n")
        assert frontmatter.endswith("---\n\n")

    def test_yaml_extract_roundtrip(self) -> None:
        """User-path: spreadsheet YAML frontmatter -> extract_yaml -> verify."""
        frontmatter = generate_basic_yaml_frontmatter("SalesData")
        full_md = frontmatter + "# Sales Report\n\n| Q1 | Q2 |\n|----|----|\n| 100 | 200 |\n"
        yaml_str, body = extract_yaml(full_md)
        assert "title: SalesData" in yaml_str
        assert "# Sales Report" in body
        assert "| Q1 | Q2 |" in body

    def test_yaml_matches_spreadsheet_converter_output_shape(self) -> None:
        """The frontmatter shape matches what ``SpreadsheetToMarkdownConverter``
        produces -- two ``---`` fences, title, aliases, blank trailing line."""
        frontmatter = generate_basic_yaml_frontmatter("Workbook1")
        lines = frontmatter.split("\n")
        assert lines[0] == "---"
        assert lines[1] == "title: Workbook1"
        assert lines[2] == "aliases:"
        assert lines[3] == "  - Workbook1"
        assert lines[4] == "---"
        assert lines[5] == ""


@pytest.mark.unit
class TestSpreadsheetImageLinkFormatting:
    """Image link formatting consumed through ``docwen_core.markdown_utils``.

    These tests verify that the shared ``format_sanitized_image_link`` utility
    produces correct wiki/markdown links with proper sanitization,
    covering the user-facing link style options available to all
    Markdown-output converters (including the spreadsheet converter).
    """

    def test_wiki_embed_link(self) -> None:
        link = format_sanitized_image_link("chart.png", style="wiki_embed")
        assert link == "![[chart.png]]"

    def test_markdown_embed_link(self) -> None:
        link = format_sanitized_image_link("chart.png", style="markdown_embed")
        assert link == "![chart.png](chart.png)"

    def test_wiki_link_strips_special_chars(self) -> None:
        link = format_sanitized_image_link(
            "chart [Q1] #final.png",
            style="wiki_embed",
        )
        assert link == "![[chart Q1 final.png]]"

    def test_markdown_link_url_encodes(self) -> None:
        link = format_sanitized_image_link("图表 2024.png", style="markdown_embed")
        assert "%E5%9B%BE%E8%A1%A8%202024.png" in link
        assert "图表 2024.png" in link

    @pytest.mark.parametrize(
        "style,expected_prefix,expected_suffix",
        [
            ("wiki_embed", "![[", "]]"),
            ("wiki_link", "[[", "]]"),
            ("markdown_embed", "![image.png](", ")"),
            ("markdown_link", "[image.png](", ")"),
        ],
    )
    def test_all_four_styles_produce_valid_links(
        self,
        style: str,
        expected_prefix: str,
        expected_suffix: str,
    ) -> None:
        link = format_sanitized_image_link("image.png", style=style)
        assert link.startswith(expected_prefix)
        assert link.endswith(expected_suffix)
