"""Inspect XLSX cells that cannot fit in the legacy BIFF8 worksheet grid."""

from __future__ import annotations

import posixpath
import re
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS = {"m": _MAIN_NS, "pr": _PACKAGE_REL_NS}
_CELL_REFERENCE_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")

LEGACY_XLS_MAX_ROWS = 65_536
LEGACY_XLS_MAX_COLUMNS = 256


@dataclass(frozen=True, slots=True)
class LegacyXlsSheetOverflow:
    """Populated cells on one sheet that exceed the BIFF8 grid."""

    name: str
    cell_count: int
    formula_count: int
    row_cell_count: int
    column_cell_count: int


@dataclass(frozen=True, slots=True)
class LegacyXlsLimitInspection:
    """Exact populated-cell overflow facts for a prospective XLS conversion."""

    out_of_bounds_cell_count: int
    out_of_bounds_formula_count: int
    out_of_bounds_row_cell_count: int
    out_of_bounds_column_cell_count: int
    affected_sheets: tuple[LegacyXlsSheetOverflow, ...]

    @property
    def has_truncation_risk(self) -> bool:
        return self.out_of_bounds_cell_count > 0


class LegacyXlsInspectionError(ValueError):
    """The XLSX package could not be inspected safely."""


def _column_number(letters: str) -> int:
    number = 0
    for character in letters.upper():
        number = number * 26 + (ord(character) - ord("A") + 1)
    return number


def _worksheet_parts(package: ZipFile) -> list[tuple[str, str]]:
    workbook = etree.fromstring(
        package.read("xl/workbook.xml"),
        parser=etree.XMLParser(resolve_entities=False, no_network=True),
    )
    relationships = etree.fromstring(
        package.read("xl/_rels/workbook.xml.rels"),
        parser=etree.XMLParser(resolve_entities=False, no_network=True),
    )
    targets = {
        relationship.get("Id", ""): relationship.get("Target", "")
        for relationship in relationships.findall("./pr:Relationship", namespaces=_NS)
        if str(relationship.get("Type", "")).endswith("/worksheet")
    }
    worksheets: list[tuple[str, str]] = []
    for sheet in workbook.findall("./m:sheets/m:sheet", namespaces=_NS):
        relationship_id = sheet.get(f"{{{_REL_NS}}}id", "")
        target = targets.get(relationship_id, "").replace("\\", "/").lstrip("/")
        if not target:
            continue
        part = (
            posixpath.normpath(target) if target.startswith("xl/") else posixpath.normpath(posixpath.join("xl", target))
        )
        worksheets.append((sheet.get("name", "worksheet"), part))
    return worksheets


def inspect_legacy_xls_limits(path: str | Path) -> LegacyXlsLimitInspection:
    """Count populated XLSX cells that XLS cannot represent.

    The scan reads only package XML and never opens or refreshes external
    workbook targets.
    """

    affected: list[LegacyXlsSheetOverflow] = []
    try:
        with ZipFile(path) as package:
            for sheet_name, part in _worksheet_parts(package):
                cell_count = 0
                formula_count = 0
                row_cell_count = 0
                column_cell_count = 0
                with package.open(part) as stream:
                    iterator = etree.iterparse(
                        stream,
                        events=("end",),
                        tag=f"{{{_MAIN_NS}}}c",
                        resolve_entities=False,
                        no_network=True,
                    )
                    for _event, cell in iterator:
                        reference = str(cell.get("r", ""))
                        match = _CELL_REFERENCE_RE.fullmatch(reference)
                        if match is not None:
                            column = _column_number(match.group(1))
                            row = int(match.group(2))
                            row_overflow = row > LEGACY_XLS_MAX_ROWS
                            column_overflow = column > LEGACY_XLS_MAX_COLUMNS
                            if row_overflow or column_overflow:
                                cell_count += 1
                                row_cell_count += int(row_overflow)
                                column_cell_count += int(column_overflow)
                                formula_count += int(cell.find(f"{{{_MAIN_NS}}}f") is not None)
                        cell.clear()
                        parent = cell.getparent()
                        if parent is not None:
                            while cell.getprevious() is not None:
                                del parent[0]
                if cell_count:
                    affected.append(
                        LegacyXlsSheetOverflow(
                            name=sheet_name,
                            cell_count=cell_count,
                            formula_count=formula_count,
                            row_cell_count=row_cell_count,
                            column_cell_count=column_cell_count,
                        )
                    )
    except (BadZipFile, KeyError, OSError, ValueError, etree.XMLSyntaxError) as exc:
        raise LegacyXlsInspectionError("The XLSX package cannot be inspected for legacy XLS limits.") from exc

    return LegacyXlsLimitInspection(
        out_of_bounds_cell_count=sum(sheet.cell_count for sheet in affected),
        out_of_bounds_formula_count=sum(sheet.formula_count for sheet in affected),
        out_of_bounds_row_cell_count=sum(sheet.row_cell_count for sheet in affected),
        out_of_bounds_column_cell_count=sum(sheet.column_cell_count for sheet in affected),
        affected_sheets=tuple(affected),
    )
