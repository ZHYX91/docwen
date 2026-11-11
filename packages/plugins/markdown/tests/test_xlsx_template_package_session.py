"""Original package ownership survives complete template save sessions."""

from __future__ import annotations

from copy import deepcopy
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.data_source import NumData, NumDataSource, NumVal
from openpyxl.workbook.defined_name import DefinedName

from docwen_plugin_markdown.xlsx_template import (
    load_template_workbook,
    register_template_clone,
    save_xlsx_workbook,
)

pytestmark = pytest.mark.contract

_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_P = "http://schemas.openxmlformats.org/package/2006/relationships"
_C = "http://schemas.openxmlformats.org/package/2006/content-types"
_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _required(root: etree._Element, path: str) -> etree._Element:
    result = root.find(path)
    assert result is not None, path
    return result


def _read(path):
    with ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def _write(path, parts):
    with ZipFile(path, "w") as package:
        for name, payload in parts.items():
            package.writestr(name, payload)


def _source(path, *, chart=False):
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Data"
    sheet["A1"] = 2
    sheet["B1"] = "=A1*3"
    sheet.sheet_format.defaultRowHeight = 26.25
    sheet.defined_names.add(DefinedName("OwnedCell", attr_text="'Data'!$A$1"))
    if chart:
        chartsheet = workbook.create_chartsheet("Plot")
        drawing = BarChart()
        drawing.add_data(Reference(sheet, min_col=1, min_row=1, max_row=1))
        chartsheet.add_chart(drawing)
    workbook.save(path)
    parts = _read(path)
    root = etree.fromstring(parts["xl/worksheets/sheet1.xml"])
    row = _required(root, f"{{{_S}}}sheetData/{{{_S}}}row")
    row.set("customHeight", "0")
    row.set("ht", "17.25")
    row.set("{urn:docwen:opaque-test}original", "keep")
    value = _required(root, f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c[@r='B1']/{{{_S}}}v")
    value.text = "6"
    parts["xl/worksheets/sheet1.xml"] = etree.tostring(root)
    parts["xl/calcChain.xml"] = f'<calcChain xmlns="{_S}"><c r="B1" i="1"/></calcChain>'.encode()
    relations = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
    etree.SubElement(
        relations,
        f"{{{_P}}}Relationship",
        Id="rIdCalcChain",
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
        Target="calcChain.xml",
    )
    parts["xl/_rels/workbook.xml.rels"] = etree.tostring(relations)
    types = etree.fromstring(parts["[Content_Types].xml"])
    etree.SubElement(
        types,
        f"{{{_C}}}Override",
        PartName="/xl/calcChain.xml",
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
    )
    parts["[Content_Types].xml"] = etree.tostring(types)
    _write(path, parts)
    return parts


def test_no_edit_keeps_every_original_member_and_formula_cache(tmp_path):
    source = tmp_path / "source.xlsx"
    before = _source(source)
    workbook = load_template_workbook(source)
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(workbook, output)
    assert _read(output) == before
    assert _read(source) == before


def test_value_edit_invalidates_caches_and_calculation_chain(tmp_path):
    source = tmp_path / "source.xlsx"
    _source(source)
    workbook = load_template_workbook(source)
    workbook["Data"]["A1"] = 7
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(workbook, output)
    parts = _read(output)
    root = etree.fromstring(parts["xl/worksheets/sheet1.xml"])
    row = _required(root, f"{{{_S}}}sheetData/{{{_S}}}row")
    assert row.get("customHeight") == "0"
    assert row.get("ht") == "17.25"
    assert row.get("{urn:docwen:opaque-test}original") == "keep"
    formula = _required(row, f"{{{_S}}}c[@r='B1']")
    assert _required(formula, f"{{{_S}}}f").text == "A1*3"
    assert formula.find(f"{{{_S}}}v") is None
    assert "xl/calcChain.xml" not in parts
    assert b"calcChain" not in parts["[Content_Types].xml"]
    assert b"calcChain" not in parts["xl/_rels/workbook.xml.rels"]
    calc = _required(etree.fromstring(parts["xl/workbook.xml"]), f"{{{_S}}}calcPr")
    assert calc.get("fullCalcOnLoad") == calc.get("forceFullCalc") == "1"


def test_chartsheet_and_scoped_names_survive_sheet_insertion(tmp_path):
    source = tmp_path / "source.xlsx"
    before = _source(source, chart=True)
    workbook = load_template_workbook(source)
    workbook.create_sheet("Added", 0)["A1"] = "new"
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(workbook, output)
    parts = _read(output)
    for name, payload in before.items():
        if name.startswith(("xl/chartsheets/", "xl/charts/", "xl/drawings/")):
            assert parts[name] == payload
    root = etree.fromstring(parts["xl/workbook.xml"])
    assert [sheet.get("name") for sheet in _required(root, f"{{{_S}}}sheets")] == ["Added", "Data", "Plot"]
    name = _required(root, f"{{{_S}}}definedNames/{{{_S}}}definedName[@name='OwnedCell']")
    assert name.get("localSheetId") == "1"
    reopened = load_workbook(output)
    assert reopened.sheetnames == ["Added", "Data", "Plot"]
    assert reopened["Added"]["A1"].value == "new"


def test_chartsheet_no_edit_is_a_package_noop(tmp_path):
    source = tmp_path / "source.xlsx"
    before = _source(source, chart=True)
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(load_template_workbook(source), output)
    assert _read(output) == before


def test_new_workbook_preserves_empty_string_and_genuine_blank(tmp_path):
    workbook = Workbook()
    workbook.worksheets[0]["A1"] = ""
    workbook.worksheets[0]["B1"] = None
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(workbook, output)
    root = etree.fromstring(_read(output)["xl/worksheets/sheet1.xml"])
    cell = _required(root, f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c[@r='A1']")
    assert cell.get("t") == "inlineStr"
    assert cell.find(f"{{{_S}}}is/{{{_S}}}t") is not None
    assert root.find(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c[@r='B1']") is None


def test_template_cannot_be_overwritten_by_its_session(tmp_path):
    source = tmp_path / "source.xlsx"
    before = _source(source)
    with pytest.raises(ValueError, match="must not overwrite"):
        save_xlsx_workbook(load_template_workbook(source), source)
    assert _read(source) == before


def test_template_clone_inherits_local_names_and_relocates_only_self_references(tmp_path):
    source = tmp_path / "source.xlsx"
    original = Workbook()
    prototype = original.worksheets[0]
    prototype.title = "Owner's Data"
    prototype["E1"] = 1
    prototype["D1"] = "=Rate"
    prototype.defined_names.add(DefinedName("Rate", attr_text="'Owner''s Data'!$E$1"))
    prototype.defined_names.add(DefinedName("Literal", attr_text='"Owner\'s Data!E1"'))
    original.save(source)
    workbook = load_template_workbook(source)
    copied = workbook.copy_worksheet(workbook.worksheets[0])
    copied.title = "Filled"
    register_template_clone(workbook, copied, workbook.worksheets[0])
    copied["E1"] = 999
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(workbook, output)
    root = etree.fromstring(_read(output)["xl/workbook.xml"])
    names = _required(root, f"{{{_S}}}definedNames")
    rate = _required(names, f"{{{_S}}}definedName[@name='Rate'][@localSheetId='1']")
    assert rate.text == "'Filled'!$E$1"
    literal = _required(names, f"{{{_S}}}definedName[@name='Literal'][@localSheetId='1']")
    assert literal.text == '"Owner\'s Data!E1"'
    original_rate = _required(names, f"{{{_S}}}definedName[@name='Rate'][@localSheetId='0']")
    assert original_rate.text == "'Owner''s Data'!$E$1"
    reopened = load_workbook(output)
    assert reopened["Filled"].defined_names["Rate"].attr_text == "'Filled'!$E$1"
    assert reopened["Filled"]["D1"].value == "=Rate"


def _source_with_chart_cache(path):
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet["A1"] = 2
    sheet["A2"] = 3
    chart = BarChart()
    chart.add_data(Reference(sheet, min_col=1, min_row=1, max_row=2))
    literal = deepcopy(chart.series[0])
    literal.val = NumDataSource(numLit=NumData(ptCount=2, pt=[NumVal(idx=0, v=41), NumVal(idx=1, v=42)]))
    chart.series.append(literal)
    sheet.add_chart(chart, "D1")
    workbook.save(path)
    parts = _read(path)
    root = etree.fromstring(parts["xl/charts/chart1.xml"])
    reference = _required(root, f".//{{{_CHART}}}numRef")
    cached = etree.SubElement(reference, f"{{{_CHART}}}numCache")
    etree.SubElement(cached, f"{{{_CHART}}}ptCount", val="2")
    for index, value in enumerate(("2", "3")):
        etree.SubElement(etree.SubElement(cached, f"{{{_CHART}}}pt", idx=str(index)), f"{{{_CHART}}}v").text = value
    parts["xl/charts/chart1.xml"] = etree.tostring(root)
    _write(path, parts)
    return parts


def test_value_change_clears_chart_reference_cache_but_preserves_literal_series(tmp_path):
    source = tmp_path / "source.xlsx"
    _source_with_chart_cache(source)
    workbook = load_template_workbook(source)
    workbook.worksheets[0]["A1"] = 999
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(workbook, output)
    chart = etree.fromstring(_read(output)["xl/charts/chart1.xml"])
    assert chart.find(f".//{{{_CHART}}}numRef/{{{_CHART}}}numCache") is None
    assert _required(chart, f".//{{{_CHART}}}numRef/{{{_CHART}}}f").text == "'Sheet'!$A$1:$A$2"
    literal = _required(chart, f".//{{{_CHART}}}numLit")
    assert [value.text for value in literal.findall(f"{{{_CHART}}}pt/{{{_CHART}}}v")] == ["41", "42"]


def test_no_edit_preserves_chart_reference_cache_bytes(tmp_path):
    source = tmp_path / "source.xlsx"
    before = _source_with_chart_cache(source)
    output = tmp_path / "output.xlsx"
    save_xlsx_workbook(load_template_workbook(source), output)
    assert _read(output) == before
