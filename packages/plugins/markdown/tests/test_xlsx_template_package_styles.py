"""The template owns existing style IDs and opaque style metadata."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, NamedStyle, PatternFill, Side

from docwen_plugin_markdown.xlsx_template import load_template_workbook, save_xlsx_workbook
from docwen_plugin_markdown.xlsx_template.styles import StyleMerger, capture_style_state

pytestmark = pytest.mark.contract

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OPAQUE = "{urn:docwen:style-preservation-test}owner"


def _bytes(workbook: Workbook) -> bytes:
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def _child(element: etree._Element, path: str) -> etree._Element:
    child = element.find(path)
    assert child is not None, path
    return child


def _int_attribute(element: etree._Element, name: str) -> int:
    value = element.get(name)
    assert isinstance(value, str), name
    return int(value)


def _merged_styles(merger: StyleMerger) -> etree._Element:
    payload = merger.to_bytes()
    assert payload is not None
    return etree.fromstring(payload)


def _part(payload: bytes, part: str = "xl/styles.xml") -> bytes:
    with ZipFile(BytesIO(payload)) as archive:
        return archive.read(part)


def _rewrite(payload: bytes, replacements: dict[str, bytes]) -> bytes:
    result = BytesIO()
    with ZipFile(BytesIO(payload)) as source, ZipFile(result, "w") as target:
        for info in source.infolist():
            target.writestr(info, replacements.get(info.filename, source.read(info.filename)))
    return result.getvalue()


def _template(*, duplicate_style: bool = False) -> bytes:
    workbook = Workbook()
    worksheet = workbook.worksheets[0]
    worksheet["A2"] = 12.5
    worksheet["A2"].number_format = '0.000 "source"'
    worksheet["B2"] = 3.5
    worksheet["B2"].number_format = '0.000 "source"'
    raw = _bytes(workbook)
    styles = etree.fromstring(_part(raw))
    formats = _child(styles, f"{{{_NS}}}numFmts")
    source_format = formats[0]
    old_id = source_format.get("numFmtId")
    source_format.set("numFmtId", "1001")
    source_format.set(_OPAQUE, "format")
    fonts = _child(styles, f"{{{_NS}}}fonts")
    fonts[0].set(_OPAQUE, "font")
    xfs = _child(styles, f"{{{_NS}}}cellXfs")
    for xf in xfs:
        if xf.get("numFmtId") == old_id:
            xf.set("numFmtId", "1001")
    xfs[1].set(_OPAQUE, "first")
    extension_list = etree.SubElement(styles, f"{{{_NS}}}extLst")
    extension = etree.SubElement(extension_list, f"{{{_NS}}}ext", uri="urn:docwen:test")
    etree.SubElement(extension, "{urn:docwen:style-preservation-test}payload").text = "keep"
    replacements = {}
    if duplicate_style:
        copied = deepcopy(xfs[1])
        copied.set(_OPAQUE, "second")
        xfs.append(copied)
        xfs.set("count", str(len(xfs)))
        sheet = etree.fromstring(_part(raw, "xl/worksheets/sheet1.xml"))
        cell = _child(sheet, f"{{{_NS}}}sheetData/{{{_NS}}}row/{{{_NS}}}c[@r='B2']")
        cell.set("s", "2")
        replacements["xl/worksheets/sheet1.xml"] = etree.tostring(sheet)
    replacements["xl/styles.xml"] = etree.tostring(styles)
    return _rewrite(raw, replacements)


def _donor_style(payload: bytes, coordinate: str = "A2") -> int:
    sheet = etree.fromstring(_part(payload, "xl/worksheets/sheet1.xml"))
    cell = _child(sheet, f"{{{_NS}}}sheetData/{{{_NS}}}row/{{{_NS}}}c[@r='{coordinate}']")
    return int(cell.get("s", "0"))


def test_unchanged_normalized_styles_keep_original_ids_and_bytes() -> None:
    source = _template()
    workbook = load_workbook(BytesIO(source))
    baseline = capture_style_state(workbook)
    donor = _bytes(workbook)
    merger = StyleMerger(_part(source), _part(donor), baseline)

    assert merger.map_style(_donor_style(donor), source_style_id=1) == 1
    assert merger.map_style(_donor_style(donor)) == 1
    assert merger.to_bytes() is None


def test_new_number_format_preserves_authored_style_metadata_and_allocates_id() -> None:
    source = _template()
    workbook = load_workbook(BytesIO(source))
    baseline = capture_style_state(workbook)
    workbook.worksheets[0]["A2"].number_format = '0.00000 "changed"'
    donor = _bytes(workbook)
    merger = StyleMerger(_part(source), _part(donor), baseline)

    result_id = merger.map_style(_donor_style(donor), source_style_id=1)
    result = _merged_styles(merger)
    original = etree.fromstring(_part(source))
    xfs = _child(result, f"{{{_NS}}}cellXfs")
    assert result_id == 2
    assert xfs[2].get(_OPAQUE) == "first"
    assert xfs[2].get("numFmtId") == "1002"
    assert [etree.tostring(node) for node in list(xfs)[:2]] == [
        etree.tostring(node) for node in _child(original, f"{{{_NS}}}cellXfs")
    ]
    formats = _child(result, f"{{{_NS}}}numFmts")
    assert [(node.get("numFmtId"), node.get("formatCode")) for node in formats] == [
        ("1001", '0.000 "source"'),
        ("1002", '0.00000 "changed"'),
    ]
    assert formats[0].get(_OPAQUE) == "format"
    assert _child(result, f"{{{_NS}}}fonts")[0].get(_OPAQUE) == "font"
    assert etree.tostring(_child(result, f"{{{_NS}}}extLst")) == etree.tostring(_child(original, f"{{{_NS}}}extLst"))


def test_new_style_components_reference_appended_records() -> None:
    source = _template()
    workbook = load_workbook(BytesIO(source))
    baseline = capture_style_state(workbook)
    cell = workbook.worksheets[0]["A2"]
    cell.font = Font(name="Arial", bold=True, color="A00000")
    cell.fill = PatternFill("solid", fgColor="F0E0D0")
    cell.border = Border(left=Side(style="double", color="001122"))
    donor = _bytes(workbook)
    merger = StyleMerger(_part(source), _part(donor), baseline)
    result_id = merger.map_style(_donor_style(donor), source_style_id=1)
    styles = _merged_styles(merger)
    xf = _child(styles, f"{{{_NS}}}cellXfs")[result_id]

    font = _child(styles, f"{{{_NS}}}fonts")[_int_attribute(xf, "fontId")]
    assert _child(font, f"{{{_NS}}}name").get("val") == "Arial"
    assert _child(font, f"{{{_NS}}}b").get("val") == "1"
    fill = _child(styles, f"{{{_NS}}}fills")[_int_attribute(xf, "fillId")]
    assert _child(fill, f"{{{_NS}}}patternFill").get("patternType") == "solid"
    border = _child(styles, f"{{{_NS}}}borders")[_int_attribute(xf, "borderId")]
    assert _child(border, f"{{{_NS}}}left").get("style") == "double"
    assert xf.get("numFmtId") == "1001"


def test_duplicate_normalized_styles_keep_distinct_authored_lineage() -> None:
    source = _template(duplicate_style=True)
    workbook = load_workbook(BytesIO(source))
    baseline = capture_style_state(workbook)
    workbook.worksheets[0]["A2"].number_format = "0.00%"
    workbook.worksheets[0]["B2"].number_format = "0.00%"
    donor = _bytes(workbook)
    merger = StyleMerger(_part(source), _part(donor), baseline)

    first = merger.map_style(_donor_style(donor, "A2"), source_style_id=1)
    second = merger.map_style(_donor_style(donor, "B2"), source_style_id=2)
    assert first != second
    xfs = _child(_merged_styles(merger), f"{{{_NS}}}cellXfs")
    assert xfs[first].get(_OPAQUE) == "first"
    assert xfs[second].get(_OPAQUE) == "second"
    assert xfs[first].get("numFmtId") == xfs[second].get("numFmtId") == "10"


def test_existing_duplicate_style_ids_are_not_collapsed() -> None:
    source = _template(duplicate_style=True)
    workbook = load_workbook(BytesIO(source))
    baseline = capture_style_state(workbook)
    donor = _bytes(workbook)
    merger = StyleMerger(_part(source), _part(donor), baseline)
    assert merger.map_style(1) == 1
    assert merger.map_style(2) == 2
    assert merger.to_bytes() is None


def test_new_cell_style_reuses_original_custom_format_identifier() -> None:
    source = _template()
    workbook = load_workbook(BytesIO(source))
    baseline = capture_style_state(workbook)
    workbook.worksheets[0]["A2"].font = Font(name="Courier New", italic=True)
    donor = _bytes(workbook)
    merger = StyleMerger(_part(source), _part(donor), baseline)
    style_id = merger.map_style(_donor_style(donor))
    styles = _merged_styles(merger)
    assert len(_child(styles, f"{{{_NS}}}numFmts")) == 1
    assert _child(styles, f"{{{_NS}}}cellXfs")[style_id].get("numFmtId") == "1001"


def test_new_named_style_keeps_original_definitions_and_adds_its_reference() -> None:
    source = _template()
    workbook = load_workbook(BytesIO(source))
    baseline = capture_style_state(workbook)
    named = NamedStyle(name="Task quantity", number_format='0.0000 "quantity"', font=Font(bold=True))
    workbook.add_named_style(named)
    workbook.worksheets[0]["A2"].style = "Task quantity"
    donor = _bytes(workbook)
    merger = StyleMerger(_part(source), _part(donor), baseline)
    style_id = merger.map_style(_donor_style(donor), source_style_id=1)
    styles = _merged_styles(merger)
    xf = _child(styles, f"{{{_NS}}}cellXfs")[style_id]
    named_records = _child(styles, f"{{{_NS}}}cellStyles")
    assert named_records[0].get("name") == "Normal"
    added = next(node for node in named_records if node.get("name") == "Task quantity")
    assert added.get("xfId") == xf.get("xfId")
    assert _int_attribute(added, "xfId") < len(_child(styles, f"{{{_NS}}}cellStyleXfs"))


def test_missing_donor_style_is_rejected() -> None:
    source = _template()
    workbook = load_workbook(BytesIO(source))
    baseline = capture_style_state(workbook)
    donor = _bytes(workbook)
    merger = StyleMerger(_part(source), _part(donor), baseline)
    with pytest.raises(ValueError, match="missing cell style"):
        merger.map_style(10000)


def _unstyled_template() -> bytes:
    workbook = Workbook()
    workbook.worksheets[0]["A1"] = "source"
    workbook.worksheets[0]["B2"] = 12
    original = _bytes(workbook)
    relations = etree.fromstring(_part(original, "xl/_rels/workbook.xml.rels"))
    for relation in list(relations):
        if relation.get("Type", "").endswith("/styles"):
            relations.remove(relation)
    content_types = etree.fromstring(_part(original, "[Content_Types].xml"))
    for entry in list(content_types):
        if entry.get("PartName") == "/xl/styles.xml":
            content_types.remove(entry)
    output = BytesIO()
    replacements = {
        "xl/_rels/workbook.xml.rels": etree.tostring(relations),
        "[Content_Types].xml": etree.tostring(content_types),
    }
    with ZipFile(BytesIO(original)) as source, ZipFile(output, "w") as destination:
        for info in source.infolist():
            if info.filename != "xl/styles.xml":
                destination.writestr(info, replacements.get(info.filename, source.read(info.filename)))
    return output.getvalue()


def _members(payload: bytes) -> dict[str, bytes]:
    with ZipFile(BytesIO(payload)) as archive:
        return {info.filename: archive.read(info.filename) for info in archive.infolist()}


def _has_styles_relationship(payload: bytes) -> bool:
    relations = etree.fromstring(_part(payload, "xl/_rels/workbook.xml.rels"))
    return any(relation.get("Type", "").endswith("/styles") for relation in relations)


def _has_styles_content_type(payload: bytes) -> bool:
    content_types = etree.fromstring(_part(payload, "[Content_Types].xml"))
    return any(entry.get("PartName") == "/xl/styles.xml" for entry in content_types)


def test_noop_template_without_styles_preserves_every_original_part(tmp_path: Path) -> None:
    template = tmp_path / "unstyled.xlsx"
    original = _unstyled_template()
    template.write_bytes(original)
    workbook = load_template_workbook(template)
    output = tmp_path / "output.xlsx"

    save_xlsx_workbook(workbook, output)

    assert _members(output.read_bytes()) == _members(original)
    assert template.read_bytes() == original


def test_value_edit_without_styles_does_not_create_a_styles_part(tmp_path: Path) -> None:
    template = tmp_path / "unstyled.xlsx"
    original = _unstyled_template()
    template.write_bytes(original)
    workbook = load_template_workbook(template)
    workbook.worksheets[0]["A1"] = "changed"
    output = tmp_path / "output.xlsx"

    save_xlsx_workbook(workbook, output)

    saved = output.read_bytes()
    assert "xl/styles.xml" not in _members(saved)
    assert not _has_styles_relationship(saved)
    assert not _has_styles_content_type(saved)
    reopened = load_workbook(BytesIO(saved))
    assert reopened.worksheets[0]["A1"].value == "changed"
    assert reopened.worksheets[0]["B2"].value == 12
    assert template.read_bytes() == original


def test_new_format_in_unstyled_template_introduces_complete_style_dependencies(tmp_path: Path) -> None:
    template = tmp_path / "unstyled.xlsx"
    original = _unstyled_template()
    template.write_bytes(original)
    workbook = load_template_workbook(template)
    workbook.worksheets[0]["B2"].number_format = '0.0000 "units"'
    output = tmp_path / "output.xlsx"

    save_xlsx_workbook(workbook, output)

    saved = output.read_bytes()
    assert "xl/styles.xml" in _members(saved)
    assert _has_styles_relationship(saved)
    assert _has_styles_content_type(saved)
    reopened = load_workbook(BytesIO(saved))
    assert reopened.worksheets[0]["A1"].value == "source"
    assert reopened.worksheets[0]["B2"].value == 12
    assert reopened.worksheets[0]["B2"].number_format == '0.0000 "units"'
    assert template.read_bytes() == original
