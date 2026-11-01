"""New minimal XLSX/ODS packages and independent XML facts for publication tests.

No production policy parser, formatter or cached-cell verifier is used here.
"""

from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from decimal import Decimal
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

from openpyxl import Workbook

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
FORMULA = "'[1]Remote'!A1"
ODS_MEDIA_TYPE = "application/vnd.oasis.opendocument.spreadsheet"

type SourceCells = dict[str, dict[str, tuple[str, str | None]]]
type TypedValue = tuple[str, str | Decimal | bool | None]
type TypedSheets = dict[str, list[list[TypedValue]]]


def _xml(element: ET.Element) -> bytes:
    return ET.tostring(element, encoding="utf-8", xml_declaration=True)


def _worksheets(parts: dict[str, bytes]) -> list[tuple[str, str]]:
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    relationships = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    by_id = {
        item.attrib["Id"]: item.attrib["Target"] for item in relationships if item.attrib["Type"] == REL + "/worksheet"
    }
    result = []
    for sheet in workbook.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet"):
        target = by_id[sheet.attrib[f"{{{REL}}}id"]]
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        result.append((sheet.attrib["name"], part))
    return result


def write_external_xlsx(path: Path, cells: SourceCells) -> bytes:
    workbook = Workbook()
    default_sheet = workbook.active
    assert default_sheet is not None
    workbook.remove(default_sheet)
    for name, entries in cells.items():
        worksheet = workbook.create_sheet(name)
        for coordinate in entries:
            worksheet[coordinate] = "=" + FORMULA
    workbook.save(path)
    workbook.close()
    with ZipFile(path) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    for name, part in _worksheets(parts):
        root = ET.fromstring(parts[part])
        for cell in root.findall(f".//{{{MAIN}}}c"):
            kind, value = cells[name][cell.attrib["r"]]
            cell.set("t", kind)
            for cache in cell.findall(f"{{{MAIN}}}v"):
                cell.remove(cache)
            if value is not None:
                ET.SubElement(cell, f"{{{MAIN}}}v").text = value
        parts[part] = _xml(root)
    workbook_xml = ET.fromstring(parts["xl/workbook.xml"])
    refs = ET.Element(f"{{{MAIN}}}externalReferences")
    following = {f"{{{MAIN}}}definedNames", f"{{{MAIN}}}calcPr"}
    before = next((index for index, child in enumerate(workbook_xml) if child.tag in following), len(workbook_xml))
    workbook_xml.insert(before, refs)
    ET.SubElement(refs, f"{{{MAIN}}}externalReference", {f"{{{REL}}}id": "rIdExternal"})
    parts["xl/workbook.xml"] = _xml(workbook_xml)
    relationships = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    ET.SubElement(
        relationships,
        f"{{{PKG_REL}}}Relationship",
        {"Id": "rIdExternal", "Type": REL + "/externalLink", "Target": "externalLinks/externalLink1.xml"},
    )
    parts["xl/_rels/workbook.xml.rels"] = _xml(relationships)
    types = ET.fromstring(parts["[Content_Types].xml"])
    ET.SubElement(
        types,
        f"{{{CONTENT_TYPES}}}Override",
        {
            "PartName": "/xl/externalLinks/externalLink1.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
        },
    )
    parts["[Content_Types].xml"] = _xml(types)
    external = ET.Element(f"{{{MAIN}}}externalLink")
    book = ET.SubElement(external, f"{{{MAIN}}}externalBook", {f"{{{REL}}}id": "rIdPath"})
    names = ET.SubElement(book, f"{{{MAIN}}}sheetNames")
    ET.SubElement(names, f"{{{MAIN}}}sheetName", {"val": "Remote"})
    parts["xl/externalLinks/externalLink1.xml"] = _xml(external)
    links = ET.Element(f"{{{PKG_REL}}}Relationships")
    ET.SubElement(
        links,
        f"{{{PKG_REL}}}Relationship",
        {"Id": "rIdPath", "Type": REL + "/externalLinkPath", "Target": "missing-remote.xlsx", "TargetMode": "External"},
    )
    parts["xl/externalLinks/_rels/externalLink1.xml.rels"] = _xml(links)
    with ZipFile(path, "w", ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)
    assert_external_graph(path)
    return path.read_bytes()


def assert_external_graph(path: Path) -> None:
    with ZipFile(path) as package:
        workbook = ET.fromstring(package.read("xl/workbook.xml"))
        relationships = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        reference = workbook.find(f"{{{MAIN}}}externalReferences/{{{MAIN}}}externalReference")
        assert reference is not None
        relation = next(item for item in relationships if item.attrib["Id"] == reference.attrib[f"{{{REL}}}id"])
        assert relation.attrib["Type"] == REL + "/externalLink"
        part = posixpath.normpath("xl/" + relation.attrib["Target"])
        external = ET.fromstring(package.read(part))
        book = external.find(f"{{{MAIN}}}externalBook")
        assert book is not None
        assert [item.attrib["val"] for item in book.findall(f"{{{MAIN}}}sheetNames/{{{MAIN}}}sheetName")] == ["Remote"]
        rel_part = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        links = ET.fromstring(package.read(rel_part))
        link = next(item for item in links if item.attrib["Id"] == book.attrib[f"{{{REL}}}id"])
        assert link.attrib == {
            "Id": "rIdPath",
            "Type": REL + "/externalLinkPath",
            "Target": "missing-remote.xlsx",
            "TargetMode": "External",
        }
        types = ET.fromstring(package.read("[Content_Types].xml"))
        assert any(item.get("PartName") == "/" + part for item in types)


def xlsx_cache_facts(path: Path) -> dict[tuple[str, str], tuple[str, bool, str | None, str | None]]:
    with ZipFile(path) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    facts = {}
    for name, part in _worksheets(parts):
        for cell in ET.fromstring(parts[part]).findall(f".//{{{MAIN}}}c"):
            value = cell.find(f"{{{MAIN}}}v")
            formula = cell.find(f"{{{MAIN}}}f")
            facts[name, cell.attrib["r"]] = (
                cell.attrib.get("t", "n"),
                value is not None,
                (value.text or "") if value is not None else None,
                formula.text if formula is not None else None,
            )
    return facts


def assert_prepared_xlsx(path: Path, cells: SourceCells) -> None:
    expected = {
        (sheet, coordinate): (kind, value is not None, value, None)
        for sheet, entries in cells.items()
        for coordinate, (kind, value) in entries.items()
    }
    assert xlsx_cache_facts(path) == expected
    with ZipFile(path) as package:
        assert not any(name.startswith("xl/externalLinks/") for name in package.namelist())
        workbook = ET.fromstring(package.read("xl/workbook.xml"))
        assert workbook.find(f"{{{MAIN}}}externalReferences") is None
        relationships = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        assert all(item.get("Type") != REL + "/externalLink" for item in relationships)


def ods_cell(kind: str, value: str = "", *, repeat: int = 1) -> str:
    repeated = f' table:number-columns-repeated="{repeat}"' if repeat != 1 else ""
    if kind == "blank":
        return f"<table:table-cell{repeated}/>"
    if kind == "covered":
        return f"<table:covered-table-cell{repeated}/>"
    if kind == "string":
        attrs = ' office:value-type="string"'
    elif kind == "number":
        attrs = ' office:value-type="float" office:value=' + quoteattr(value)
    elif kind == "boolean":
        attrs = ' office:value-type="boolean" office:boolean-value=' + quoteattr(value)
    else:
        raise AssertionError(kind)
    return f"<table:table-cell{attrs}{repeated}><text:p>{escape(value)}</text:p></table:table-cell>"


def ods_row(cells_xml: str, *, repeat: int = 1) -> str:
    repeated = f' table:number-rows-repeated="{repeat}"' if repeat != 1 else ""
    return f"<table:table-row{repeated}>{cells_xml}</table:table-row>"


def write_ods(path: Path, sheets: list[tuple[str, str]]) -> bytes:
    tables = "".join(f"<table:table table:name={quoteattr(name)}>{rows}</table:table>" for name, rows in sheets)
    content = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.2">'
        f"<office:body><office:spreadsheet>{tables}</office:spreadsheet></office:body></office:document-content>"
    ).encode()
    manifest = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">'
        f'<manifest:file-entry manifest:full-path="/" manifest:media-type="{ODS_MEDIA_TYPE}"/>'
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        "</manifest:manifest>"
    ).encode()
    with ZipFile(path, "w") as package:
        package.writestr("mimetype", ODS_MEDIA_TYPE, compress_type=ZIP_STORED)
        package.writestr("content.xml", content, compress_type=ZIP_DEFLATED)
        package.writestr("META-INF/manifest.xml", manifest, compress_type=ZIP_DEFLATED)
    return path.read_bytes()


def ods_typed_matrix(path: Path) -> TypedSheets:
    """Independent small-fixture ODF consumer: preserve type, sheet and position."""
    with ZipFile(path) as package:
        assert package.read("mimetype").decode() == ODS_MEDIA_TYPE
        root = ET.fromstring(package.read("content.xml"))
    result: TypedSheets = {}
    for table in root.findall(f"{{{OFFICE}}}body/{{{OFFICE}}}spreadsheet/{{{TABLE}}}table"):
        matrix = []
        for row in table.findall(f"{{{TABLE}}}table-row"):
            values: list[TypedValue] = []
            for cell in row:
                kind = cell.get(f"{{{OFFICE}}}value-type")
                if cell.tag == f"{{{TABLE}}}covered-table-cell":
                    fact: TypedValue = ("covered", None)
                elif kind == "string":
                    text = cell.get(f"{{{OFFICE}}}string-value")
                    if text is None:
                        text = "\n".join("".join(paragraph.itertext()) for paragraph in cell.findall(f"{{{TEXT}}}p"))
                    fact = ("string", text)
                elif kind == "float":
                    fact = ("number", Decimal(cell.attrib[f"{{{OFFICE}}}value"]))
                elif kind == "boolean":
                    raw = cell.attrib[f"{{{OFFICE}}}boolean-value"]
                    assert raw in {"true", "false", "1", "0"}
                    fact = ("boolean", raw in {"true", "1"})
                elif kind is None:
                    scalar = cell.get(f"{{{OFFICE}}}value")
                    fact = ("untyped-scalar", scalar) if scalar is not None else ("blank", None)
                else:
                    raise AssertionError(kind)
                repeated = int(cell.get(f"{{{TABLE}}}number-columns-repeated", "1"))
                assert 1 <= repeated <= 10
                values.extend([fact] * repeated)
            repeated_rows = int(row.get(f"{{{TABLE}}}number-rows-repeated", "1"))
            assert 1 <= repeated_rows <= 10
            matrix.extend([list(values) for _ in range(repeated_rows)])
        name = table.attrib[f"{{{TABLE}}}name"]
        assert name not in result
        result[name] = matrix
    return result
