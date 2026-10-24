"""Inspect stored worksheet formula cells without materializing sparse grids."""

from __future__ import annotations

import posixpath
from collections.abc import Callable, Sequence
from zipfile import ZipFile

from lxml import etree
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PACKAGE = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def visit_formula_caches(
    path: str,
    visit: Callable[[str, str, bool], None],
    *,
    sheet_names: Sequence[str] | None = None,
    cancel_check: Callable[[], None] | None = None,
) -> None:
    """Visit actual formula coordinates and whether XML proves an empty string.

    Only t=str with an explicit empty v is known empty, not a missing cache.
    Worksheet relationships exclude chartsheets and external relationships.
    Streams close even when the visitor or cancellation raises.
    """
    selected = None if sheet_names is None else set(sheet_names)
    with ZipFile(path) as package:
        parser = etree.XMLParser(resolve_entities=False, no_network=True)
        workbook = etree.fromstring(package.read("xl/workbook.xml"), parser)
        relationships = etree.fromstring(package.read("xl/_rels/workbook.xml.rels"), parser)
        targets = {
            relation.get("Id"): relation.get("Target", "")
            for relation in relationships.findall(f"{_PACKAGE}Relationship")
            if str(relation.get("Type", "")).endswith("/worksheet") and relation.get("TargetMode") != "External"
        }
        for sheet in workbook.findall(f"{_MAIN}sheets/{_MAIN}sheet"):
            if cancel_check:
                cancel_check()
            name = str(sheet.get("name", ""))
            if selected is not None and name not in selected:
                continue
            target = str(targets.get(sheet.get(f"{_REL}id"), "")).replace("\\", "/")
            if not target:
                continue
            part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else f"xl/{target}")
            with package.open(part) as stream:
                iterator = etree.iterparse(
                    stream,
                    events=("start", "end"),
                    tag=(f"{_MAIN}c", f"{_MAIN}row"),
                    resolve_entities=False,
                    no_network=True,
                )
                row_number = column_number = visited = 0
                for event, element in iterator:
                    if event == "start":
                        if element.tag == f"{_MAIN}row":
                            row_number = int(element.get("r", row_number + 1))
                            column_number = 0
                        continue
                    if element.tag == f"{_MAIN}c":
                        visited += 1
                        if cancel_check and visited % 1000 == 0:
                            cancel_check()
                        coordinate = element.get("r")
                        if coordinate:
                            _, column_number = coordinate_to_tuple(coordinate)
                        else:
                            column_number += 1
                            coordinate = f"{get_column_letter(column_number)}{row_number}"
                        if element.find(f"{_MAIN}f") is not None:
                            value = element.find(f"{_MAIN}v")
                            known_empty = element.get("t") == "str" and value is not None and not value.text
                            visit(name, coordinate, known_empty)
                    element.clear()
                    parent = element.getparent()
                    if parent is not None:
                        while element.getprevious() is not None:
                            del parent[0]
                if cancel_check:
                    cancel_check()
