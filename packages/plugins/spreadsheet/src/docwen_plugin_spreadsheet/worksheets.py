"""Ordinary worksheet identity shared by native spreadsheet consumers."""

from __future__ import annotations

import posixpath
from collections.abc import Iterator
from typing import Any
from zipfile import ZipFile

from lxml import etree
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PACKAGE = "{http://schemas.openxmlformats.org/package/2006/relationships}"


class NoOrdinaryWorksheetsError(ValueError):
    """A chart-only workbook has no worksheet output to publish."""

    diagnostic_code = "SPREADSHEET-NO-WORKSHEETS"


def ordinary_worksheets(workbook: Any) -> tuple[Any, ...]:
    """Retain physical workbook order using worksheet type, not sheet names."""
    sheets = tuple(sheet for sheet in workbook.worksheets if isinstance(sheet, (Worksheet, ReadOnlyWorksheet)))
    if not sheets:
        raise NoOrdinaryWorksheetsError("The workbook contains no ordinary worksheets to export.")
    return sheets


def ordinary_worksheet_parts(package: ZipFile) -> Iterator[tuple[str, str]]:
    """Resolve the actual internal worksheet relationship graph in order."""
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    workbook = etree.fromstring(package.read("xl/workbook.xml"), parser)
    relationships = etree.fromstring(package.read("xl/_rels/workbook.xml.rels"), parser)
    targets = {
        relation.get("Id"): relation.get("Target", "")
        for relation in relationships.findall(f"{_PACKAGE}Relationship")
        if str(relation.get("Type", "")).endswith("/worksheet") and relation.get("TargetMode") != "External"
    }
    for sheet in workbook.findall(f"{_MAIN}sheets/{_MAIN}sheet"):
        target = str(targets.get(sheet.get(f"{_REL}id"), "")).replace("\\", "/")
        if target:
            part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else f"xl/{target}")
            yield str(sheet.get("name", "")), part
