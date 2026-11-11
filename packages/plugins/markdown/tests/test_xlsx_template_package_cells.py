"""Worksheet edits preserve XML properties outside their authored scope."""

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.worksheet.worksheet import Worksheet

from docwen_plugin_markdown.xlsx_template import load_template_workbook, save_xlsx_workbook
from docwen_plugin_markdown.xlsx_template.state import SheetState
from docwen_plugin_markdown.xlsx_template.worksheet import patch_worksheet

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
Q = "{" + NS + "}"
PART = "xl/worksheets/sheet1.xml"
pytestmark = pytest.mark.unit


class Styles:
    def __init__(self):
        self.calls = []

    def map_style(self, value, *, source_style_id=None):
        self.calls.append((value, source_style_id))
        return value


def active_sheet(workbook: Workbook) -> Worksheet:
    sheet = workbook.active
    assert isinstance(sheet, Worksheet)
    return sheet


def required(root: etree._Element, path: str) -> etree._Element:
    node = root.find(path)
    assert node is not None, path
    return node


def package(workbook: Workbook) -> dict[str, bytes]:
    stream = BytesIO()
    workbook.save(stream)
    with ZipFile(stream) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def original_workbook():
    workbook = Workbook()
    sheet = active_sheet(workbook)
    sheet["A1"] = "old value"
    sheet["A2"] = "=1+2"
    sheet["A5"] = "wrapping text"
    sheet.row_dimensions[5].height = 17.25
    sheet.sheet_format.defaultRowHeight = 26.25
    parts = package(workbook)
    root = etree.fromstring(parts[PART])
    row = required(root, f"{Q}sheetData/{Q}row[@r='5']")
    row.attrib.pop("customHeight")
    row.set("{urn:preserved}setting", "original")
    cell = required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='A1']")
    cell.set("cm", "4")
    extension = etree.SubElement(cell, Q + "extLst")
    etree.SubElement(extension, "{urn:preserved}payload").text = "untouched"
    source = BytesIO()
    with ZipFile(source, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, etree.tostring(root) if name == PART else content)
    loaded = load_workbook(BytesIO(source.getvalue()))
    return loaded, root, SheetState.capture(active_sheet(loaded))


def apply(workbook, root, baseline):
    final = SheetState.capture(active_sheet(workbook))
    donor = etree.fromstring(package(workbook)[PART])
    styles = Styles()
    result = patch_worksheet(root, donor, active_sheet(workbook), baseline, styles, final_state=final)
    return result, styles


def test_value_fill_preserves_row_state_cell_metadata_and_formula():
    workbook, root, baseline = original_workbook()
    formula = etree.tostring(required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='A2']"))
    active_sheet(workbook)["A1"] = "new value"
    result, styles = apply(workbook, root, baseline)
    assert result.changed and result.calculation_changed
    assert styles.calls == []
    row = required(root, f"{Q}sheetData/{Q}row[@r='5']")
    assert row.get("ht") == "17.25"
    assert row.get("customHeight") is None
    assert row.get("{urn:preserved}setting") == "original"
    cell = required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='A1']")
    assert cell.get("cm") == "4"
    assert required(cell, f"{Q}is/{Q}t").text == "new value"
    assert required(cell, f"{Q}extLst/{{urn:preserved}}payload").text == "untouched"
    assert etree.tostring(required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='A2']")) == formula


def test_no_change_does_not_rewrite_xml_or_infer_fixed_height():
    workbook, root, baseline = original_workbook()
    before = etree.tostring(root)
    result, styles = apply(workbook, root, baseline)
    assert not result.changed
    assert not result.calculation_changed
    assert etree.tostring(root) == before
    assert styles.calls == []


def test_explicit_empty_is_distinct_from_cleared_value():
    workbook, root, baseline = original_workbook()
    active_sheet(workbook)["A1"] = ""
    active_sheet(workbook)["A5"] = None
    apply(workbook, root, baseline)
    empty = required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='A1']")
    cleared = required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='A5']")
    assert empty.get("t") == "inlineStr"
    assert empty.find(f"{Q}is/{Q}t") is not None
    assert required(empty, f"{Q}is/{Q}t").get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert cleared.find(Q + "is") is None
    assert cleared.find(Q + "v") is None
    assert cleared.get("t") is None


def test_authored_height_change_sets_fixed_height_preserving_unknowns():
    workbook, root, baseline = original_workbook()
    active_sheet(workbook).row_dimensions[5].height = 42
    result, _ = apply(workbook, root, baseline)
    row = required(root, f"{Q}sheetData/{Q}row[@r='5']")
    assert row.get("ht") == "42.0"
    assert row.get("customHeight") == "1"
    assert row.get("{urn:preserved}setting") == "original"
    assert result.changed and not result.calculation_changed


def test_single_column_resize_splits_inherited_range_and_keeps_metadata():
    workbook, root, baseline = original_workbook()
    cols = etree.Element(Q + "cols")
    etree.SubElement(
        cols, Q + "col", {"min": "6", "max": "16384", "width": "9", "style": "0", "{urn:preserved}setting": "original"}
    )
    root.insert(list(root).index(required(root, Q + "sheetData")), cols)
    active_sheet(workbook).column_dimensions["H"].width = 32
    apply(workbook, root, baseline)
    columns = root.findall(f"{Q}cols/{Q}col")
    assert [(item.get("min"), item.get("max"), item.get("width")) for item in columns] == [
        ("6", "7", "9"),
        ("8", "8", "32.0"),
        ("9", "16384", "9"),
    ]
    assert all(item.get("{urn:preserved}setting") == "original" for item in columns)


def test_comment_and_hyperlink_changes_are_reported_without_replacing_cell():
    workbook, root, baseline = original_workbook()
    active_sheet(workbook)["A1"].comment = Comment("new note", "Author")
    active_sheet(workbook)["A1"].hyperlink = "https://example.test/resource"
    result, _ = apply(workbook, root, baseline)
    assert result.comment_changes == {"A1"}
    assert result.hyperlink_changes == {"A1"}
    assert not result.calculation_changed
    assert required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='A1']").get("cm") == "4"


def test_default_height_edit_keeps_unmodeled_sheet_format_attributes():
    workbook, root, baseline = original_workbook()
    required(root, Q + "sheetFormatPr").set("{urn:preserved}setting", "original")
    active_sheet(workbook).sheet_format.defaultRowHeight = 31
    apply(workbook, root, baseline)
    item = required(root, Q + "sheetFormatPr")
    assert item.get("defaultRowHeight") == "31"
    assert item.get("{urn:preserved}setting") == "original"


def test_merging_new_range_preserves_existing_merge_metadata():
    workbook, root, baseline = original_workbook()
    merges = etree.SubElement(root, Q + "mergeCells", count="1")
    original = etree.SubElement(merges, Q + "mergeCell", {"ref": "D8:E8", "{urn:preserved}setting": "original"})
    active_sheet(workbook).merge_cells("D8:E8")
    baseline = SheetState.capture(active_sheet(workbook))
    original_bytes = etree.tostring(deepcopy(original))
    active_sheet(workbook).merge_cells("A8:B8")
    result, _ = apply(workbook, root, baseline)
    assert result.calculation_changed
    assert required(root, Q + "mergeCells").get("count") == "2"
    assert etree.tostring(required(root, f"{Q}mergeCells/{Q}mergeCell[@ref='D8:E8']")) == original_bytes


def test_filling_inside_original_dimension_does_not_shrink_its_bounds():
    workbook, root, baseline = original_workbook()
    required(root, Q + "dimension").set("ref", "A1:Z100")
    active_sheet(workbook)["B3"] = "new value"
    apply(workbook, root, baseline)
    assert required(root, Q + "dimension").get("ref") == "A1:Z100"


def test_new_column_properties_preserve_source_xml_comments():
    workbook, root, baseline = original_workbook()
    root.insert(0, etree.Comment("source annotation"))
    active_sheet(workbook).column_dimensions["F"].width = 20
    apply(workbook, root, baseline)
    assert root[0].text == "source annotation"
    assert required(root, f"{Q}cols/{Q}col[@min='6']").get("width") == "20.0"


def shared_formula_package() -> bytes:
    workbook = Workbook()
    sheet = active_sheet(workbook)
    for row in range(2, 5):
        sheet[f"A{row}"] = f"=B{row}*2"
        sheet[f"B{row}"] = row
        sheet[f"C{row}"] = f"=D{row}+10"
        sheet[f"D{row}"] = row
    parts = package(workbook)
    root = etree.fromstring(parts[PART])
    for column, identity in (("A", "0"), ("C", "1")):
        for row in range(2, 5):
            cell = required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='{column}{row}']")
            formula = required(cell, Q + "f")
            formula.set("t", "shared")
            formula.set("si", identity)
            if row == 2:
                formula.set("ref", f"{column}2:{column}4")
            else:
                formula.text = None
            required(cell, Q + "v").text = str(row * 2 if column == "A" else row + 10)
    output = BytesIO()
    with ZipFile(output, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, etree.tostring(root) if name == PART else content)
    return output.getvalue()


@pytest.mark.parametrize("overwritten", ["A2", "A3"])
def test_shared_formula_group_survives_overwritten_master_or_follower(tmp_path: Path, overwritten: str):
    source = tmp_path / "shared.xlsx"
    destination = tmp_path / "filled.xlsx"
    source.write_bytes(shared_formula_package())
    with ZipFile(source) as archive:
        original = etree.fromstring(archive.read(PART))
    unrelated_formulas = {
        f"C{row}": etree.tostring(required(original, f"{Q}sheetData/{Q}row/{Q}c[@r='C{row}']/{Q}f"))
        for row in range(2, 5)
    }
    workbook = load_template_workbook(source)
    active_sheet(workbook)[overwritten] = "replacement"
    save_xlsx_workbook(workbook, destination)
    with ZipFile(destination) as archive:
        actual = etree.fromstring(archive.read(PART))
        workbook_xml = etree.fromstring(archive.read("xl/workbook.xml"))
    for row in range(2, 5):
        address = f"A{row}"
        cell = required(actual, f"{Q}sheetData/{Q}row/{Q}c[@r='{address}']")
        formula = cell.find(Q + "f")
        if address == overwritten:
            assert formula is None
            assert required(cell, f"{Q}is/{Q}t").text == "replacement"
        else:
            assert formula is not None
            assert formula.text == f"B{row}*2"
            assert all(formula.get(attribute) is None for attribute in ("t", "ref", "si"))
            assert cell.find(Q + "v") is None
        other = required(actual, f"{Q}sheetData/{Q}row/{Q}c[@r='C{row}']")
        assert etree.tostring(required(other, Q + "f")) == unrelated_formulas[f"C{row}"]
        assert other.find(Q + "v") is None
    assert required(workbook_xml, Q + "calcPr").get("fullCalcOnLoad") == "1"
    reopened = load_workbook(destination)
    for row in range(2, 5):
        assert active_sheet(reopened)[f"A{row}"].value == ("replacement" if f"A{row}" == overwritten else f"=B{row}*2")
        assert active_sheet(reopened)[f"C{row}"].value == f"=D{row}+10"


def test_untouched_shared_formula_groups_keep_original_cache_and_xml(tmp_path: Path):
    source = tmp_path / "shared.xlsx"
    destination = tmp_path / "unchanged.xlsx"
    source.write_bytes(shared_formula_package())
    workbook = load_template_workbook(source)
    save_xlsx_workbook(workbook, destination)
    with ZipFile(source) as before, ZipFile(destination) as after:
        assert before.read(PART) == after.read(PART)


@pytest.mark.parametrize("invalid", ["missing-master", "outside-range"])
def test_edited_shared_formula_with_ambiguous_source_fails(invalid: str):
    payload = shared_formula_package()
    workbook = load_workbook(BytesIO(payload))
    baseline = SheetState.capture(active_sheet(workbook))
    with ZipFile(BytesIO(payload)) as archive:
        root = etree.fromstring(archive.read(PART))
    master = required(root, f"{Q}sheetData/{Q}row/{Q}c[@r='A2']/{Q}f")
    if invalid == "missing-master":
        master.text = None
    else:
        master.set("ref", "A2:A3")
    active_sheet(workbook)["A2"] = "replacement"
    with pytest.raises(ValueError, match="shared formula"):
        apply(workbook, root, baseline)
