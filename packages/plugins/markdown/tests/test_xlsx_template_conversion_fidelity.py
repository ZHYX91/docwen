"""Template exports preserve native settings outside their authored changes."""

from __future__ import annotations

from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Alignment, Font
from openpyxl.styles.differential import DifferentialStyle

from docwen_plugin_markdown.to_spreadsheet.converter import MdToXlsxConverter

from .conftest import make_context

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _template(path: Path, *, formula: bool = False) -> None:
    workbook = Workbook()
    fixed = workbook.active
    assert fixed is not None
    fixed.title = "Fixed"
    fixed["A1"] = "{{title}}"
    fixed["A5"] = "first\nsecond"
    fixed["A5"].alignment = Alignment(wrap_text=True)
    fixed["A5"].font = Font(bold=True, italic=True)
    fixed.sheet_format.defaultRowHeight = 26.25
    fixed.sheet_format.customHeight = True
    fixed.conditional_formatting.add(
        "A5",
        Rule(type="expression", formula=["1=1"], dxf=DifferentialStyle(font=Font(bold=False, italic=False))),
    )
    automatic = workbook.copy_worksheet(fixed)
    automatic.title = "Automatic"
    automatic["A1"] = "Untouched"
    if formula:
        fixed["A2"] = "=LEN(A1)"
    workbook.save(path)
    with ZipFile(path) as package:
        parts = {info.filename: package.read(info) for info in package.infolist()}
    for part, attributes in (
        ("xl/worksheets/sheet1.xml", {"customHeight": "1"}),
        ("xl/worksheets/sheet2.xml", {"ht": "17.25"}),
    ):
        root = etree.fromstring(parts[part])
        row = root.find(f"{{{_NS}}}sheetData/{{{_NS}}}row[@r='5']")
        assert row is not None
        for name, value in attributes.items():
            row.set(name, value)
        if formula and part.endswith("sheet1.xml"):
            cache = root.find(f"{{{_NS}}}sheetData/{{{_NS}}}row/{{{_NS}}}c[@r='A2']/{{{_NS}}}v")
            assert cache is not None
            cache.text = "9"
        parts[part] = etree.tostring(root)
    styles = etree.fromstring(parts["xl/styles.xml"])
    font = styles.find(f"{{{_NS}}}dxfs/{{{_NS}}}dxf/{{{_NS}}}font")
    assert font is not None
    for attribute in ("b", "i"):
        etree.SubElement(font, f"{{{_NS}}}{attribute}", val="0")
    extensions = etree.SubElement(styles, f"{{{_NS}}}extLst")
    extension = etree.SubElement(extensions, f"{{{_NS}}}ext", uri="urn:docwen-test:keep")
    etree.SubElement(extension, "{urn:docwen-test:keep}setting", value="retain")
    parts["xl/styles.xml"] = etree.tostring(styles)
    with ZipFile(path, "w") as package:
        for name, data in parts.items():
            package.writestr(name, data)


def _convert(tmp_path: Path, template: Path) -> Path:
    source = tmp_path / "source.md"
    source.write_text("---\ntitle: Filled\n---\n", encoding="utf-8")
    context, _ = make_context(str(source), target_format="xlsx", options={"template_name": str(template)})
    result = MdToXlsxConverter().convert(context)
    assert result.success, result.error
    return Path(result.artifacts[0].staging_path)


@pytest.mark.contract
def test_template_export_preserves_height_states_and_conditional_format_instructions(tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    _template(template)
    before_hash = sha256(template.read_bytes()).digest()
    output = _convert(tmp_path, template)
    assert sha256(template.read_bytes()).digest() == before_hash
    with ZipFile(template) as original, ZipFile(output) as generated:
        assert generated.read("xl/styles.xml") == original.read("xl/styles.xml")
        assert generated.read("xl/worksheets/sheet2.xml") == original.read("xl/worksheets/sheet2.xml")
        fixed = etree.fromstring(generated.read("xl/worksheets/sheet1.xml"))
        row = fixed.find(f"{{{_NS}}}sheetData/{{{_NS}}}row[@r='5']")
        assert row is not None
        assert row.get("ht") is None
        assert row.get("customHeight") == "1"
    workbook = load_workbook(output)
    assert workbook["Fixed"]["A1"].value == "Filled"
    assert workbook["Automatic"]["A1"].value == "Untouched"


@pytest.mark.contract
def test_template_export_invalidates_stale_formula_cache_after_filling(tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    _template(template, formula=True)
    output = _convert(tmp_path, template)
    with ZipFile(output) as generated:
        root = etree.fromstring(generated.read("xl/worksheets/sheet1.xml"))
        cell = root.find(f"{{{_NS}}}sheetData/{{{_NS}}}row/{{{_NS}}}c[@r='A2']")
        assert cell is not None
        formula = cell.find(f"{{{_NS}}}f")
        assert formula is not None and formula.text == "LEN(A1)"
        cache = cell.find(f"{{{_NS}}}v")
        assert cache is None or not cache.text
        workbook = etree.fromstring(generated.read("xl/workbook.xml"))
        calculation = workbook.find(f"{{{_NS}}}calcPr")
        assert calculation is not None and calculation.get("fullCalcOnLoad") == "1"
