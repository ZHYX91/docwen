from __future__ import annotations

from pathlib import Path
from typing import cast
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

import pytest
from lxml import etree

from docwen_plugin_spreadsheet.format_conversion.ods_grid_compaction import (
    compact_generated_ods_grid,
)

pytestmark = pytest.mark.contract


def _write_boundary_xlsx(path: Path) -> None:
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as package:
        package.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        )
        package.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
   Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        package.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/><sheetData/>
</worksheet>""",
        )


def _write_repeated_ods(path: Path) -> bytes:
    content = b"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 office:version="1.3">
 <office:body><office:spreadsheet>
  <table:table table:name="Sheet1">
   <table:table-column table:number-columns-repeated="10"/>
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>kept</text:p></table:table-cell>
    <table:table-cell table:style-name="keep-style"/>
    <table:table-cell table:style-name="trailing" table:number-columns-repeated="8"/>
   </table:table-row>
   <table:table-row table:number-rows-repeated="5">
    <table:table-cell table:style-name="row-style" table:number-columns-repeated="10"/>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>"""
    with ZipFile(path, "w") as package:
        package.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet", compress_type=ZIP_STORED)
        package.writestr("content.xml", content, compress_type=ZIP_DEFLATED)
        package.writestr("unchanged.bin", b"unchanged", compress_type=ZIP_DEFLATED)
    return content


def test_generated_ods_grid_is_bounded_to_producing_xlsx_dimensions(
    tmp_path: Path,
) -> None:
    ods = tmp_path / "generated.ods"
    boundary = tmp_path / "boundary.xlsx"
    _write_repeated_ods(ods)
    _write_boundary_xlsx(boundary)

    result = compact_generated_ods_grid(ods, boundary)

    assert result.changed is True
    assert result.removed_repeated_rows == 4
    assert result.removed_repeated_columns > 0
    assert result.trimmed_sheets[0].sheet_name == "Sheet1"
    assert result.trimmed_sheets[0].max_row == 2
    assert result.trimmed_sheets[0].max_column == 2
    with ZipFile(ods) as package:
        assert package.testzip() is None
        assert package.read("unchanged.bin") == b"unchanged"
        root = etree.fromstring(package.read("content.xml"))
    namespaces = {
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    }
    assert root.xpath("string(.//text:p)", namespaces=namespaces) == "kept"
    rows = cast(list[etree._Element], root.xpath(".//table:table-row", namespaces=namespaces))
    assert len(rows) == 2
    assert rows[1].get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-rows-repeated") is None
    for row in rows:
        cells = cast(list[etree._Element], row.xpath("./table:table-cell", namespaces=namespaces))
        logical_columns = sum(
            int(
                cell.get(
                    "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-repeated",
                    "1",
                )
            )
            for cell in cells
        )
        assert logical_columns == 2


def test_generated_ods_grid_keeps_semantic_content_beyond_cell_dimension(
    tmp_path: Path,
) -> None:
    ods = tmp_path / "generated-with-drawing.ods"
    boundary = tmp_path / "boundary.xlsx"
    _write_boundary_xlsx(boundary)
    content = b"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
 office:version="1.3">
 <office:body><office:spreadsheet>
  <table:table table:name="Sheet1">
   <table:table-row>
    <table:table-cell table:number-columns-repeated="2"/>
    <table:table-cell><draw:frame draw:name="Chart anchor"/></table:table-cell>
    <table:table-cell table:style-name="trailing" table:number-columns-repeated="7"/>
   </table:table-row>
   <table:table-row table:number-rows-repeated="9">
    <table:table-cell table:number-columns-repeated="10"/>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>"""
    with ZipFile(ods, "w") as package:
        package.writestr(
            "mimetype",
            "application/vnd.oasis.opendocument.spreadsheet",
            compress_type=ZIP_STORED,
        )
        package.writestr("content.xml", content, compress_type=ZIP_DEFLATED)

    result = compact_generated_ods_grid(ods, boundary)

    assert result.trimmed_sheets[0].max_column == 3
    with ZipFile(ods) as package:
        root = etree.fromstring(package.read("content.xml"))
    namespaces = {
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    }
    frames = cast(list[etree._Element], root.xpath(".//draw:frame", namespaces=namespaces))
    assert len(frames) == 1
    rows = cast(list[etree._Element], root.xpath(".//table:table-row", namespaces=namespaces))
    first_row = rows[0]
    logical_columns = sum(
        int(
            cell.get(
                "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-repeated",
                "1",
            )
        )
        for cell in cast(list[etree._Element], first_row.xpath("./table:table-cell", namespaces=namespaces))
    )
    assert logical_columns == 3
