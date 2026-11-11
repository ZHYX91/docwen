"""Saved package references follow clone ownership and individual link edits."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table

from docwen_plugin_markdown.xlsx_template import (
    load_template_workbook,
    register_template_clone,
    save_xlsx_workbook,
)

pytestmark = pytest.mark.contract

_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _parts(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def test_two_clones_retarget_local_names_and_formulas_to_their_own_tables(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Prototype"
    sheet.append(["Label", "Value"])
    sheet.append(["original", 4])
    sheet.add_table(Table(displayName="Items", ref="A1:B2"))
    sheet["D1"] = "=Rate"
    sheet["D2"] = "=SUM(Items[Value])"
    sheet.defined_names.add(DefinedName("Rate", attr_text="SUM(Items[Value])"))
    sheet.defined_names.add(DefinedName("Combined", attr_text="SUM(Items[Value])+'Prototype'!$B$2"))
    sheet.defined_names.add(DefinedName("Qualified", attr_text="SUM('Prototype'!Items[Value])"))
    sheet.defined_names.add(DefinedName("Literal", attr_text='"Items[Value] Prototype!B2"'))
    sheet.defined_names.add(DefinedName("External", attr_text="SUM('[remote.xlsx]Prototype'!Items[Value])"))
    workbook.defined_names.add(DefinedName("GlobalItems", attr_text="SUM(Items[Value])"))
    workbook.save(source)
    before = source.read_bytes()

    edited = load_template_workbook(source)
    for title, value in (("First", 999), ("Second", 27)):
        clone = edited.copy_worksheet(edited["Prototype"])
        clone.title = title
        register_template_clone(edited, clone, edited["Prototype"])
        clone["B2"] = value
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(edited, output)

    parts = _parts(output)
    names = etree.fromstring(parts["xl/workbook.xml"]).find(f"{{{_S}}}definedNames")
    assert names is not None
    actual = {(node.get("localSheetId"), node.get("name")): node.text for node in names}
    reopened = load_workbook(output)
    for index, title, table_name, value in (
        (0, "Prototype", "Items", 4),
        (1, "First", "Items_copy1", 999),
        (2, "Second", "Items_copy2", 27),
    ):
        scope = str(index)
        assert actual[scope, "Rate"] == f"SUM({table_name}[Value])"
        assert actual[scope, "Combined"] == f"SUM({table_name}[Value])+'{title}'!$B$2"
        assert actual[scope, "Qualified"] == f"SUM('{title}'!{table_name}[Value])"
        assert actual[scope, "Literal"] == '"Items[Value] Prototype!B2"'
        assert actual[scope, "External"] == "SUM('[remote.xlsx]Prototype'!Items[Value])"
        final_sheet = reopened[title]
        assert list(final_sheet.tables) == [table_name]
        assert final_sheet["B2"].value == value
        assert final_sheet["D1"].value == "=Rate"
        assert final_sheet["D2"].value == f"=SUM({table_name}[Value])"
        assert final_sheet.defined_names["Rate"].attr_text == f"SUM({table_name}[Value])"
    assert actual[None, "GlobalItems"] == "SUM(Items[Value])"
    assert source.read_bytes() == before


def _range_link_source(path: Path, *, internal: bool, merged: bool = False) -> bytes:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Prototype"
    for row in range(1, 4):
        for column in range(1, 4):
            sheet.cell(row, column, f"value-{row}-{column}")
    if merged:
        sheet.merge_cells("A1:C1")
    sheet["A1"].hyperlink = Hyperlink(
        ref="A1",
        location="Prototype!C3" if internal else None,
        target=None if internal else "https://example.test/original",
        tooltip="Original tooltip",
        display="Original display",
    )
    workbook.save(path)
    parts = _parts(path)
    root = etree.fromstring(parts["xl/worksheets/sheet1.xml"])
    link = root.find(f"{{{_S}}}hyperlinks/{{{_S}}}hyperlink")
    assert link is not None
    link.set("ref", "A1:C1" if merged else "A1:C3")
    parts["xl/worksheets/sheet1.xml"] = etree.tostring(root)
    with ZipFile(path, "w") as package:
        for name, payload in parts.items():
            package.writestr(name, payload)
    return path.read_bytes()


@pytest.mark.parametrize("clone", [False, True])
@pytest.mark.parametrize("internal", [False, True])
@pytest.mark.parametrize("edit", ["clear_one", "replace_one", "clear_all"])
def test_range_hyperlink_edits_preserve_only_untouched_cells(
    tmp_path: Path, clone: bool, internal: bool, edit: str
) -> None:
    source = tmp_path / "source.xlsx"
    before = _range_link_source(source, internal=internal)
    workbook = load_template_workbook(source)
    sheet = workbook["Prototype"]
    if clone:
        sheet = workbook.copy_worksheet(sheet)
        sheet.title = "Filled"
        register_template_clone(workbook, sheet, workbook["Prototype"])
    changed = {"B2"} if edit != "clear_all" else {f"{column}{row}" for column in "ABC" for row in range(1, 4)}
    for coordinate in changed:
        sheet[coordinate].hyperlink = "https://example.test/replacement" if edit == "replace_one" else None
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(workbook, output)

    final = load_workbook(output)[sheet.title]
    for row in range(1, 4):
        for column in "ABC":
            coordinate = f"{column}{row}"
            link = final[coordinate].hyperlink
            if coordinate in changed:
                if edit == "replace_one":
                    assert link is not None and link.target == "https://example.test/replacement"
                else:
                    assert link is None
            else:
                assert link is not None
                assert link.location == ("Prototype!C3" if internal else None)
                assert link.target == (None if internal else "https://example.test/original")
                assert (link.tooltip, link.display) == ("Original tooltip", "Original display")
    parts = _parts(output)
    part = "xl/worksheets/sheet1_copy1.xml" if clone else "xl/worksheets/sheet1.xml"
    root = etree.fromstring(parts[part])
    covered: set[str] = set()
    for link in root.findall(f"{{{_S}}}hyperlinks/{{{_S}}}hyperlink"):
        reference = link.get("ref")
        assert reference is not None
        min_col, min_row, max_col, max_row = range_boundaries(reference)
        assert min_col is not None and min_row is not None and max_col is not None and max_row is not None
        cells = {
            f"{get_column_letter(column)}{row}"
            for row in range(min_row, max_row + 1)
            for column in range(min_col, max_col + 1)
        }
        assert not covered.intersection(cells), "Saved hyperlink definitions must not overlap"
        covered.update(cells)
    all_cells = {f"{column}{row}" for column in "ABC" for row in range(1, 4)}
    assert covered == (all_cells if edit == "replace_one" else all_cells - changed)
    if clone:
        assert parts["xl/worksheets/sheet1.xml"] == _parts(source)["xl/worksheets/sheet1.xml"]
    assert source.read_bytes() == before


@pytest.mark.parametrize("internal", [False, True])
def test_clearing_merged_range_anchor_does_not_restore_its_link(tmp_path: Path, internal: bool) -> None:
    source = tmp_path / "source.xlsx"
    before = _range_link_source(source, internal=internal, merged=True)
    workbook = load_template_workbook(source)
    workbook["Prototype"]["A1"].hyperlink = None
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(workbook, output)

    reopened = load_workbook(output)
    assert reopened["Prototype"]["A1"].hyperlink is None
    assert str(reopened["Prototype"].merged_cells) == "A1:C1"
    assert source.read_bytes() == before
