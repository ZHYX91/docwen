"""Structured table projection retains the selected template's package data."""

from __future__ import annotations

import hashlib
import posixpath
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as WorksheetImage
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName
from PIL import Image

from docwen_plugin_markdown.structured_clipboard_xlsx import (
    create_table_sheet,
    prepare_projection_workbook,
    write_string_matrix,
)
from docwen_plugin_markdown.xlsx_template import save_xlsx_workbook

pytestmark = pytest.mark.contract

SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EXTENSION_NS = "urn:docwen:test:template-layout"
NS = {"s": SHEET_NS, "layout": EXTENSION_NS}


def _worksheet_parts(package: ZipFile) -> dict[str, str]:
    workbook = etree.fromstring(package.read("xl/workbook.xml"))
    relations = etree.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relation.get("Id"): posixpath.normpath(
            relation.get("Target", "").lstrip("/")
            if relation.get("Target", "").startswith("/")
            else posixpath.join("xl", relation.get("Target", ""))
        )
        for relation in relations
        if relation.get("Type", "").endswith("/worksheet")
    }
    return {
        sheet.get("name", ""): targets[sheet.get(f"{{{OFFICE_REL_NS}}}id")]
        for sheet in workbook.findall("s:sheets/s:sheet", NS)
    }


def _template_with_original_layout(path: Path, *, image_payload: bytes | None = None) -> dict[str, bytes]:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Prototype"
    sheet.sheet_format.defaultRowHeight = 26.25
    sheet.sheet_format.customHeight = True
    sheet["A1"] = "Original merged title"
    sheet.merge_cells("A1:B1")
    sheet["A2"] = "Original detail"
    sheet["A5"] = "Untouched template text"
    sheet["A5"].font = Font(name="Arial", bold=True, size=12)
    sheet.row_dimensions[5].height = 17.25
    sheet["A6"] = "Inherited row height"
    sheet["D5"] = "Outside the table"
    sheet.merge_cells("D5:E5")
    sheet["D6"] = "Reference"
    sheet["D6"].hyperlink = "https://example.com/template-reference"
    sheet.freeze_panes = "C3"
    sheet.print_area = "A1:E6"
    if image_payload is not None:
        for anchor in ("A8", "D8"):
            sheet.add_image(WorksheetImage(BytesIO(image_payload)), anchor)
    reference = workbook.create_sheet("Reference")
    reference["A1"] = "Untouched reference"
    reference.sheet_state = "hidden"
    reference.defined_names.add(DefinedName("LocalReference", attr_text="'Reference'!$A$1"))
    workbook.save(path)
    workbook.close()

    with ZipFile(path) as package:
        parts = {info.filename: package.read(info.filename) for info in package.infolist()}
    root = etree.fromstring(parts["xl/worksheets/sheet1.xml"])
    fifth = root.find("s:sheetData/s:row[@r='5']", NS)
    sixth = root.find("s:sheetData/s:row[@r='6']", NS)
    assert fifth is not None and sixth is not None
    fifth.attrib.pop("customHeight", "")
    sixth.set("customHeight", "1")
    extensions = etree.SubElement(root, f"{{{SHEET_NS}}}extLst")
    extension = etree.SubElement(extensions, f"{{{SHEET_NS}}}ext", uri=EXTENSION_NS)
    etree.SubElement(extension, f"{{{EXTENSION_NS}}}layout", {"version": "1", "mode": "authored"})
    parts["xl/worksheets/sheet1.xml"] = etree.tostring(root, encoding="utf-8", xml_declaration=True)
    with ZipFile(path, "w") as package:
        for part, payload in parts.items():
            package.writestr(part, payload)
    return parts


def test_projection_keeps_original_sheets_and_clones_unmodified_layout(tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    original = _template_with_original_layout(template)
    original_hash = hashlib.sha256(template.read_bytes()).digest()
    projection = prepare_projection_workbook(str(template))
    table = create_table_sheet(
        projection,
        table_id="table-1",
        requested_name="Extracted table",
        row_count=2,
        column_count=2,
    )
    write_string_matrix(table, ("Name", "Value"), [("Filled value", "")])
    write_string_matrix(projection.order_sheet, ("Kind", "Sheet"), [("table", table.title)])
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(projection.workbook, output)

    with ZipFile(output) as package:
        parts = _worksheet_parts(package)
        assert package.read(parts["Prototype"]) == original["xl/worksheets/sheet1.xml"]
        assert package.read(parts["Reference"]) == original["xl/worksheets/sheet2.xml"]
        root = etree.fromstring(package.read(parts[table.title]))
        fifth = root.find("s:sheetData/s:row[@r='5']", NS)
        sixth = root.find("s:sheetData/s:row[@r='6']", NS)
        assert fifth is not None and sixth is not None
        assert fifth.get("ht") == "17.25" and fifth.get("customHeight") is None
        assert sixth.get("ht") is None and sixth.get("customHeight") == "1"
        sheet_format = root.find("s:sheetFormatPr", NS)
        assert sheet_format is not None and sheet_format.get("defaultRowHeight") == "26.25"
        layout = root.find("s:extLst/s:ext/layout:layout", NS)
        assert layout is not None and layout.get("mode") == "authored"
        merged = {item.get("ref") for item in root.findall("s:mergeCells/s:mergeCell", NS)}
        assert merged == {"D5:E5"}
        empty = root.find("s:sheetData/s:row[@r='2']/s:c[@r='B2']", NS)
        assert empty is not None and empty.get("t") == "inlineStr"
        assert empty.find("s:is/s:t", NS) is not None
        workbook_root = etree.fromstring(package.read("xl/workbook.xml"))
        defined_names = workbook_root.findall("s:definedNames/s:definedName", NS)
        reference_name = next(item for item in defined_names if item.get("name") == "LocalReference")
        assert reference_name.get("localSheetId") == "3"
        assert reference_name.text == "'Reference'!$A$1"

    result = load_workbook(output)
    assert result.sheetnames == ["Document Order", "Table Semantics", "Prototype", "Reference", "Extracted table"]
    assert result["Prototype"]["A1"].value == "Original merged title"
    assert str(result["Prototype"].print_area) == "'Prototype'!$A$1:$E$6"
    assert result["Reference"].sheet_state == "hidden"
    assert result["Reference"].defined_names["LocalReference"].attr_text == "'Reference'!$A$1"
    assert result["Extracted table"]["A2"].value == "Filled value"
    assert result["Extracted table"]["A5"].value == "Untouched template text"
    assert result["Extracted table"]["A5"].font.bold is True
    assert result["Extracted table"]["D6"].hyperlink.target == "https://example.com/template-reference"
    assert result["Extracted table"].freeze_panes == "C3"
    assert str(result["Extracted table"].print_area) == "'Extracted table'!$A$1:$E$6"
    assert result["Document Order"]["B2"].value == "Extracted table"
    result.close()
    assert hashlib.sha256(template.read_bytes()).digest() == original_hash


def test_projection_retains_existing_media_and_reuses_new_repeated_images(tmp_path: Path) -> None:
    old_image = BytesIO()
    new_image = BytesIO()
    Image.new("RGB", (3, 2), (50, 100, 150)).save(old_image, format="PNG")
    Image.new("RGB", (3, 2), (150, 100, 50)).save(new_image, format="PNG")
    template = tmp_path / "images-template.xlsx"
    original = _template_with_original_layout(template, image_payload=old_image.getvalue())
    original_hash = hashlib.sha256(template.read_bytes()).digest()
    original_media = {name: payload for name, payload in original.items() if name.startswith("xl/media/")}
    assert len(original_media) == 2
    projection = prepare_projection_workbook(str(template))
    table = create_table_sheet(
        projection,
        table_id="table-1",
        requested_name="Images",
        row_count=1,
        column_count=1,
    )
    table["A1"] = "Filled table"
    for anchor in ("C1", "D1"):
        table.add_image(WorksheetImage(BytesIO(new_image.getvalue())), anchor)
    output = tmp_path / "images-output.xlsx"
    save_xlsx_workbook(projection.workbook, output)

    drawing_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    with ZipFile(output) as package:
        media = {name: package.read(name) for name in package.namelist() if name.startswith("xl/media/")}
        assert len(media) == 3
        assert all(media[name] == payload for name, payload in original_media.items())
        assert sum(payload == new_image.getvalue() for payload in media.values()) == 1
        parts = _worksheet_parts(package)
        assert package.read(parts["Prototype"]) == original["xl/worksheets/sheet1.xml"]
        original_drawing = "xl/drawings/drawing1.xml"
        assert package.read(original_drawing) == original[original_drawing]
        clone = etree.fromstring(package.read(parts["Images"]))
        drawing = clone.find("s:drawing", NS)
        assert drawing is not None
        relation_path = posixpath.join(
            posixpath.dirname(parts["Images"]), "_rels", posixpath.basename(parts["Images"]) + ".rels"
        )
        relations = etree.fromstring(package.read(relation_path))
        relation = next(item for item in relations if item.get("Id") == drawing.get(f"{{{OFFICE_REL_NS}}}id"))
        target = relation.get("Target", "")
        drawing_path = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join(posixpath.dirname(parts["Images"]), target))
        )
        drawing_root = etree.fromstring(package.read(drawing_path))
        pictures = drawing_root.findall(f".//{{{drawing_ns}}}pic")
        identifiers = [item.get("id") for item in drawing_root.findall(f".//{{{drawing_ns}}}cNvPr")]
        assert len(pictures) == 4
        assert len(identifiers) == len(set(identifiers)) == 4
    assert hashlib.sha256(template.read_bytes()).digest() == original_hash


@pytest.mark.parametrize("with_template", [False, True])
def test_projection_writes_empty_and_formula_looking_strings_with_shared_writer(
    tmp_path: Path, with_template: bool
) -> None:
    template = tmp_path / "template.xlsx"
    if with_template:
        _template_with_original_layout(template)
    projection = prepare_projection_workbook(str(template) if with_template else None)
    table = create_table_sheet(
        projection,
        table_id="table-1",
        requested_name="Values",
        row_count=2,
        column_count=3,
    )
    write_string_matrix(table, ("Empty", "Literal", "Text"), [("", "=SUM(A1:A2)", "line one\nline two")])
    output = tmp_path / "strings.xlsx"
    save_xlsx_workbook(projection.workbook, output)

    with ZipFile(output) as package:
        root = etree.fromstring(package.read(_worksheet_parts(package)["Values"]))
        empty = root.find("s:sheetData/s:row[@r='2']/s:c[@r='A2']", NS)
        literal = root.find("s:sheetData/s:row[@r='2']/s:c[@r='B2']", NS)
        assert empty is not None and empty.find("s:is/s:t", NS) is not None
        assert literal is not None and literal.find("s:f", NS) is None

    result = load_workbook(output)
    assert result["Values"]["A2"].value == ""
    assert result["Values"]["B2"].value == "=SUM(A1:A2)"
    assert result["Values"]["B2"].data_type == "s"
    assert result["Values"]["C2"].value == "line one\nline two"
    assert result["Values"]["C2"].number_format == "@"
    result.close()
