"""Typed OOXML worksheet inventory and allocation checks before native loading."""

from __future__ import annotations

import posixpath
import re
from collections.abc import Callable, Iterator
from pathlib import Path
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree
from xml.parsers import expat
from zipfile import ZipFile

from docwen_core.spreadsheet_projection import (
    MAX_SPREADSHEET_PROJECTION_CELLS,
    NoOrdinaryWorksheetsError,
    ProjectionBudget,
    SpreadsheetProjectionLimitError,
)

_MAIN_URI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_MAIN = "{" + _MAIN_URI + "}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PACKAGE = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_COORDINATE = re.compile(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)\Z")
_MAX_PACKAGE_MEMBERS = 10_000
_MAX_MEMBER_BYTES = 64 * 1024 * 1024
_MAX_PACKAGE_BYTES = 256 * 1024 * 1024
_MAX_METADATA_BYTES = 4 * 1024 * 1024


class _PackageTreeBuilder(ElementTree.TreeBuilder):
    def doctype(self, name: str, pubid: str | None, system: str | None) -> None:
        raise ValueError("Spreadsheet package XML must not declare a document type.")


def _parse(payload: bytes) -> ElementTree.Element:
    return ElementTree.fromstring(payload, parser=ElementTree.XMLParser(target=_PackageTreeBuilder()))


def _check_package(package: ZipFile) -> None:
    entries = package.infolist()
    if len(entries) > _MAX_PACKAGE_MEMBERS:
        raise SpreadsheetProjectionLimitError("Spreadsheet package exceeds its member-count budget.")
    names: set[str] = set()
    total = 0
    for entry in entries:
        name = entry.orig_filename
        if (
            not name
            or name != entry.filename
            or name.casefold() in names
            or "\\" in name
            or name.startswith("/")
            or any(segment in {".", ".."} for segment in name.split("/"))
            or posixpath.normpath(name.rstrip("/")) != name.rstrip("/")
        ):
            raise ValueError("Spreadsheet package contains ambiguous part names.")
        names.add(name.casefold())
        total += entry.file_size
        if entry.file_size > _MAX_MEMBER_BYTES or total > _MAX_PACKAGE_BYTES:
            raise SpreadsheetProjectionLimitError("Spreadsheet package exceeds its uncompressed-byte budget.")


def _metadata(package: ZipFile, part: str) -> ElementTree.Element:
    if package.getinfo(part).file_size > _MAX_METADATA_BYTES:
        raise SpreadsheetProjectionLimitError("Spreadsheet metadata exceeds its byte budget.")
    with package.open(part) as stream:
        payload = stream.read(_MAX_METADATA_BYTES + 1)
    if len(payload) > _MAX_METADATA_BYTES:
        raise SpreadsheetProjectionLimitError("Spreadsheet metadata exceeds its byte budget.")
    return _parse(payload)


def ordinary_worksheet_parts(package: ZipFile) -> Iterator[tuple[str, str]]:
    """Resolve internal worksheet relationships in physical workbook order."""
    _check_package(package)
    workbook = _metadata(package, "xl/workbook.xml")
    relationships = _metadata(package, "xl/_rels/workbook.xml.rels")
    if workbook.tag != f"{_MAIN}workbook" or relationships.tag != f"{_PACKAGE}Relationships":
        raise ValueError("Spreadsheet package has invalid workbook metadata.")
    targets: dict[str, str] = {}
    identifiers: set[str] = set()
    for relation in relationships.findall(f"{_PACKAGE}Relationship"):
        identifier = relation.get("Id", "")
        if not identifier or identifier in identifiers:
            raise ValueError("Spreadsheet workbook relationships have duplicate or missing identities.")
        identifiers.add(identifier)
        if relation.get("Type") == _REL[1:-1] + "/worksheet" and relation.get("TargetMode") != "External":
            if relation.get("TargetMode") not in {None, "Internal"}:
                raise ValueError("Worksheet relationship has an invalid target mode.")
            targets[identifier] = relation.get("Target", "")
    for sheet in workbook.findall(f"{_MAIN}sheets/{_MAIN}sheet"):
        target = targets.get(sheet.get(f"{_REL}id", ""))
        if target is None:
            continue
        parsed = urlsplit(target)
        if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment or "\\" in target:
            raise ValueError("Worksheet relationship must reference an internal package part.")
        path = unquote(parsed.path)
        part = posixpath.normpath(path.lstrip("/") if path.startswith("/") else f"xl/{path}")
        if not path or part in {".", ".."} or part.startswith("../"):
            raise ValueError("Worksheet relationship is outside the package.")
        yield str(sheet.get("name", "")), part


def _coordinate(value: str) -> tuple[int, int]:
    match = _COORDINATE.fullmatch(value)
    if match is None:
        raise ValueError("Worksheet merge has no finite cell rectangle.")
    column = 0
    for character in match.group(1).upper():
        column = column * 26 + ord(character) - ord("A") + 1
    row = int(match.group(2))
    if column > 16384 or row > 1048576:
        raise ValueError("Worksheet merge exceeds native row or column bounds.")
    return row, column


def _merge_area(reference: str) -> int:
    bounds = reference.split(":")
    if len(bounds) not in {1, 2}:
        raise ValueError("Worksheet merge has no finite cell rectangle.")
    min_row, min_column = _coordinate(bounds[0])
    max_row, max_column = _coordinate(bounds[-1])
    if min_row > max_row or min_column > max_column:
        raise ValueError("Worksheet merge has reversed bounds.")
    return (max_row - min_row + 1) * (max_column - min_column + 1)


def preflight_worksheet_allocations(
    path: str | Path, *, cancel_check: Callable[[], None], limit: int = MAX_SPREADSHEET_PROJECTION_CELLS
) -> None:
    """Bound stored cells and merge coverage without materializing XML trees.

    The independent budgets are cumulative across ordinary worksheets. Distant
    cells cost one stored cell; only delimited output plans reserve rectangles.
    Package member and uncompressed-byte bounds complement the cell budgets;
    library load/save remain synchronous.
    """
    cells = ProjectionBudget(limit=limit)
    covered = ProjectionBudget(limit=limit)
    found = False
    scanned_bytes = 0
    with ZipFile(path) as package:
        for name, part in ordinary_worksheet_parts(package):
            found = True
            cancel_check()
            parser = expat.ParserCreate(namespace_separator="}")
            visited = 0

            def start(tag: str, attributes: dict[str, str], sheet_name: str = name) -> None:
                nonlocal visited
                if visited == 0 and tag != _MAIN_URI + "}worksheet":
                    raise ValueError("Worksheet relationship does not reference a worksheet document.")
                visited += 1
                if visited % 1000 == 0:
                    cancel_check()
                if tag == _MAIN_URI + "}c":
                    cells.reserve(1, location=sheet_name)
                elif tag == _MAIN_URI + "}mergeCell":
                    covered.reserve(_merge_area(attributes.get("ref", "")), location=sheet_name)

            def reject_doctype(*_args: object) -> None:
                raise ValueError("Spreadsheet package XML must not declare a document type.")

            parser.StartElementHandler = start
            parser.StartDoctypeDeclHandler = reject_doctype
            member_bytes = 0
            with package.open(part) as stream:
                while chunk := stream.read(64 * 1024):
                    cancel_check()
                    member_bytes += len(chunk)
                    scanned_bytes += len(chunk)
                    if member_bytes > _MAX_MEMBER_BYTES or scanned_bytes > _MAX_PACKAGE_BYTES:
                        raise SpreadsheetProjectionLimitError("Worksheet XML exceeds its uncompressed-byte budget.")
                    parser.Parse(chunk, False)
                parser.Parse(b"", True)
            cancel_check()
    if not found:
        raise NoOrdinaryWorksheetsError("The workbook contains no ordinary worksheets to export.")
