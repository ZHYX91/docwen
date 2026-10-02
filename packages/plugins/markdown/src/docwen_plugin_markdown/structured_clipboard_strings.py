"""Preserve authored empty XLSX strings without changing their value."""

from __future__ import annotations

import os
import posixpath
import tempfile
from contextlib import suppress
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from lxml import etree

_SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_NS = "http://www.w3.org/XML/1998/namespace"


def _parse(payload: bytes) -> Any:
    return etree.fromstring(payload, parser=etree.XMLParser(resolve_entities=False, no_network=True))


def _worksheet_parts(package: ZipFile) -> dict[str, str]:
    workbook = _parse(package.read("xl/workbook.xml"))
    relations = _parse(package.read("xl/_rels/workbook.xml.rels"))
    targets: dict[str, str] = {}
    for relation in relations.findall(f"{{{_PACKAGE_REL_NS}}}Relationship"):
        if relation.get("TargetMode") == "External" or not relation.get("Type", "").endswith("/worksheet"):
            continue
        target = relation.get("Target", "")
        part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target))
        if not part.startswith("xl/worksheets/") or part not in package.namelist():
            raise ValueError("XLSX worksheet relationship is outside the generated package")
        targets[relation.get("Id", "")] = part
    return {
        sheet.get("name", ""): targets[sheet.get(f"{{{_OFFICE_REL_NS}}}id", "")]
        for sheet in workbook.findall(f"{{{_SHEET_NS}}}sheets/{{{_SHEET_NS}}}sheet")
        if sheet.get(f"{{{_OFFICE_REL_NS}}}id", "") in targets
    }


def save_workbook_preserving_empty_strings(workbook: Any, output_path: str | Path) -> None:
    """Write a real empty inline string; leave absent/blank cells absent.

    openpyxl 3.1.5 omits the inline string element for value="". Restore it
    only for cells explicitly holding that string before serialization.
    Formula cells and genuine None values never enter this set.
    """
    authored: dict[str, set[str]] = {
        sheet.title: {cell.coordinate for cell in sheet._cells.values() if cell.data_type == "s" and cell.value == ""}
        for sheet in workbook.worksheets
    }
    authored = {title: coordinates for title, coordinates in authored.items() if coordinates}
    output = Path(output_path)
    workbook.save(output)
    if not authored:
        return

    descriptor, raw_path = tempfile.mkstemp(prefix="clipboard-strings-", suffix=".xlsx", dir=output.parent)
    os.close(descriptor)
    temporary = Path(raw_path)
    try:
        with ZipFile(output) as package:
            parts = _worksheet_parts(package)
            replacements: dict[str, bytes] = {}
            for title, coordinates in authored.items():
                part = parts[title]
                root = _parse(package.read(part))
                matched: set[str] = set()
                for cell in root.findall(f"{{{_SHEET_NS}}}sheetData/{{{_SHEET_NS}}}row/{{{_SHEET_NS}}}c"):
                    coordinate = cell.get("r", "")
                    if coordinate not in coordinates:
                        continue
                    if cell.get("t") != "inlineStr" or cell.find(f"{{{_SHEET_NS}}}f") is not None:
                        raise ValueError("Authored empty XLSX cell did not serialize as a string")
                    for child in list(cell):
                        cell.remove(child)
                    inline = etree.SubElement(cell, f"{{{_SHEET_NS}}}is")
                    text = etree.SubElement(inline, f"{{{_SHEET_NS}}}t")
                    text.set(f"{{{_XML_NS}}}space", "preserve")
                    text.text = ""
                    matched.add(coordinate)
                if matched != coordinates:
                    raise ValueError("An authored empty XLSX string was omitted from the worksheet")
                replacements[part] = etree.tostring(root, encoding="utf-8", xml_declaration=True)
            with ZipFile(temporary, "w") as rewritten:
                for info in package.infolist():
                    rewritten.writestr(info, replacements.get(info.filename, package.read(info.filename)))
        os.replace(temporary, output)
    finally:
        with suppress(FileNotFoundError):
            temporary.unlink()
