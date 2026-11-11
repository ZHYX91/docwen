"""YAML scalar identity survives the complete XLSX conversion and serialization."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook

from docwen_plugin_markdown.to_spreadsheet.converter import MdToXlsxConverter

from .conftest import make_context

pytestmark = pytest.mark.contract

_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


@pytest.mark.parametrize(
    ("yaml_value", "expected", "data_type"),
    [
        ("123456789012345", 123456789012345, "n"),
        ("-123456789012345", -123456789012345, "n"),
        ("1234567890123456", "1234567890123456", "s"),
        ("-1234567890123456", "-1234567890123456", "s"),
        ("1234567890123456789", "1234567890123456789", "s"),
        ("-1234567890123456789", "-1234567890123456789", "s"),
        ('"1234567890123456789"', "1234567890123456789", "s"),
        ('"001234567890123456789"', "001234567890123456789", "s"),
        ('"-00123"', "-00123", "s"),
        ('"1.25"', "1.25", "s"),
        ("1.25", 1.25, "n"),
        ("0", 0, "n"),
        ("true", True, "b"),
        ("false", False, "b"),
        ('"=SUM(1,2)"', "=SUM(1,2)", "f"),
        ('"1/2"', "=1/2", "f"),
    ],
)
def test_yaml_scalar_uses_its_actual_cell_type(
    tmp_path: Path, yaml_value: str, expected: object, data_type: str
) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "{{value}}"
    worksheet["A1"].number_format = "0.000"
    worksheet["B1"] = "Identifier: {{value}}"
    workbook.save(template)
    before = template.read_bytes()
    source = tmp_path / "source.md"
    source.write_text(f"---\nvalue: {yaml_value}\n---\n", encoding="utf-8")
    context, _ = make_context(str(source), target_format="xlsx", options={"template_name": str(template)})

    result = MdToXlsxConverter().convert(context)

    assert result.success, result.error
    output = Path(result.artifacts[0].staging_path)
    reopened = load_workbook(output)
    try:
        cell = reopened.worksheets[0]["A1"]
        assert cell.value == expected
        assert cell.data_type == data_type
        if data_type == "s":
            assert cell.number_format == "@"
        elif data_type in {"n", "b"}:
            assert cell.number_format == "0.000"
        assert reopened.worksheets[0]["B1"].value == f"Identifier: {expected}"
    finally:
        reopened.close()
    with ZipFile(output) as package:
        root = etree.fromstring(package.read("xl/worksheets/sheet1.xml"))
        cell = root.find(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c[@r='A1']")
        assert cell is not None
        if data_type == "s":
            assert cell.get("t") == "inlineStr"
            assert cell.findtext(f"{{{_S}}}is/{{{_S}}}t") == expected
            assert cell.find(f"{{{_S}}}v") is None
        elif data_type == "f":
            assert cell.findtext(f"{{{_S}}}f") == str(expected).removeprefix("=")
        else:
            assert cell.get("t") == data_type
    assert template.read_bytes() == before
