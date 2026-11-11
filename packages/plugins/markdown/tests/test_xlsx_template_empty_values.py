"""Ordinary exports clear generated blanks without changing authored template cells."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from docwen_plugin_markdown.to_spreadsheet.converter import MdToXlsxConverter

from ._link_processing_routes_support import _link_config
from ._md_to_spreadsheet_support import _png_bytes
from .conftest import make_context

pytestmark = pytest.mark.contract

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def test_template_cleanup_preserves_original_empty_string_cells(tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    for coordinate, value in {
        "A1": "{{missing}}",
        "A2": "{{IMAGE:photo.png}}",
        "A3": "{{IMAGE:absent.png}}",
        "A4": "before{{missing}}after",
        "A5": "before{{IMAGE:photo.png}}after",
        "B1": "{{empty}}",
        "B2": "{{null_value}}",
        "B3": "{{empty_list}}",
        "B4": "{{null_value}}{{empty}}",
        "B5": "before{{null_value}}after",
        "C1": "{{↓Value}}",
        "E1": "",
        "E2": '=""',
    }.items():
        sheet[coordinate] = value
    sheet["A1"].font = Font(bold=True)
    sheet["E1"].font = Font(italic=True)
    workbook.save(template)
    workbook.close()
    # openpyxl's own save collapses an empty string. Supply the real XLSX
    # representation so the conversion must preserve an authored empty string.
    with ZipFile(template) as package:
        parts = {info.filename: package.read(info) for info in package.infolist()}
    root = etree.fromstring(parts["xl/worksheets/sheet1.xml"])
    authored = root.find("s:sheetData/s:row/s:c[@r='E1']", NS)
    assert authored is not None
    inline = etree.SubElement(authored, f"{{{NS['s']}}}is")
    etree.SubElement(inline, f"{{{NS['s']}}}t").text = ""
    original_cell = etree.tostring(authored, method="c14n")
    parts["xl/worksheets/sheet1.xml"] = etree.tostring(root)
    with ZipFile(template, "w") as package:
        for name, data in parts.items():
            package.writestr(name, data)
    original_template = template.read_bytes()
    (tmp_path / "photo.png").write_bytes(_png_bytes())
    source = tmp_path / "source.md"
    source.write_text(
        '---\nempty: ""\nnull_value: null\nempty_list: []\n---\n\n| Value |\n| --- |\n| |\n| filled |\n',
        encoding="utf-8",
    )
    context, _workspace = make_context(str(source), target_format="xlsx", options={"template_name": str(template)})

    result = MdToXlsxConverter().convert(context)

    assert result.success, result.error
    output = Path(result.artifacts[0].staging_path)
    loaded = load_workbook(output)
    try:
        sheet = loaded.active
        assert sheet is not None
        for coordinate in ("A1", "A2", "A3", "B1", "B2", "B3", "B4", "C1"):
            assert sheet[coordinate].value is None, coordinate
        assert sheet["A1"].font.bold is True
        assert sheet["A4"].value == sheet["A5"].value == sheet["B5"].value == "beforeafter"
        assert sheet["C2"].value == "filled"
        assert sheet["E1"].value == "" and sheet["E1"].font.italic is True
        assert sheet["E2"].value == '=""' and sheet["E2"].data_type == "f"
        assert len(vars(sheet)["_images"]) == 2
    finally:
        loaded.close()
    with ZipFile(output) as package:
        root = etree.fromstring(package.read("xl/worksheets/sheet1.xml"))
        authored = root.find("s:sheetData/s:row/s:c[@r='E1']", NS)
        assert authored is not None and etree.tostring(authored, method="c14n") == original_cell
        cleared = root.find("s:sheetData/s:row/s:c[@r='A1']", NS)
        assert cleared is not None and cleared.find("s:is", NS) is None
    assert template.read_bytes() == original_template


@pytest.mark.parametrize("use_template", [False, True], ids=["new-workbook", "template"])
def test_generated_empty_cells_and_removed_links_are_blank(tmp_path: Path, use_template: bool) -> None:
    options = {}
    if use_template:
        template = tmp_path / "template.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet["A2"] = "{{↓Value}}"
        sheet["B2"] = "{{↓Control}}"
        workbook.save(template)
        workbook.close()
        options["template_name"] = str(template)
    source = tmp_path / "source.md"
    source.write_text(
        "| Value | Control |\n| --- | --- |\n| | first |\n| [removed](https://example.com) | second |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source), target_format="xlsx", options=options, config_values=_link_config(markdown_mode="remove")
    )

    result = MdToXlsxConverter().convert(context)

    assert result.success, result.error
    workbook = load_workbook(result.artifacts[0].staging_path)
    try:
        sheet = workbook.active
        assert sheet is not None
        assert sheet["A2"].value is None and sheet["A3"].value is None
        assert sheet["B2"].value == "first" and sheet["B3"].value == "second"
    finally:
        workbook.close()
