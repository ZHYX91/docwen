"""Focused tests split from test_link_processing_routes.py."""

from __future__ import annotations

from ._link_processing_routes_support import (
    MdToCsvConverter,
    MdToXlsxConverter,
    Path,
    _link_config,
    csv,
    load_workbook,
    make_context,
    pytest,
)

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("target_format", ["xlsx", "csv"])
@pytest.mark.parametrize("wiki_mode", ["keep", "hyperlink"])
def test_spreadsheet_cross_boundary_wiki_pipe_keeps_outer_table_context(
    tmp_path: Path,
    target_format: str,
    wiki_mode: str,
) -> None:
    (tmp_path / "child.md").write_text("[[target.md", encoding="utf-8")
    (tmp_path / "target.md").write_text("target\n", encoding="utf-8")
    source = tmp_path / f"boundary-table-{target_format}.md"
    source.write_text(
        "| Value | Note |\n| --- | --- |\n| ![[child.md]]|Shown]] | ok |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format=target_format,
        config_values=_link_config(wiki_mode=wiki_mode),
    )

    converter = MdToXlsxConverter() if target_format == "xlsx" else MdToCsvConverter()
    result = converter.convert(context)

    assert result.success is True, result.error
    output = Path(result.artifacts[0].staging_path)
    if target_format == "xlsx":
        workbook = load_workbook(output)
        try:
            sheet = workbook.active
            assert sheet is not None
            assert sheet.max_column == 2
            row = [sheet["A2"].value, sheet["B2"].value]
        finally:
            workbook.close()
    else:
        with output.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        row = rows[1]
        assert len(row) == 2
    assert row == ["[[target.md|Shown]]", "ok"]
