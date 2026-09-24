"""Identify explicit empty-string formula caches lost by openpyxl's None coercion."""

from __future__ import annotations

import posixpath
from collections.abc import Callable
from zipfile import ZipFile

from lxml import etree

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PACKAGE = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def empty_string_caches(path: str, *, cancel_check: Callable[[], None] | None = None) -> set[tuple[str, str]]:
    """Only t=str with an explicit empty v is known empty, not a missing cache.

    Shared-string/inline-string empty values already survive the ordinary reader.
    No external relationship is opened and the workbook is never modified.
    """
    result: set[tuple[str, str]] = set()
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
            target = str(targets.get(sheet.get(f"{_REL}id"), "")).replace("\\", "/")
            if not target:
                continue
            part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else f"xl/{target}")
            with package.open(part) as stream:
                iterator = etree.iterparse(
                    stream, events=("end",), tag=f"{_MAIN}c", resolve_entities=False, no_network=True
                )
                for index, (_, cell) in enumerate(iterator, 1):
                    if cancel_check and index % 1000 == 0:
                        cancel_check()
                    value = cell.find(f"{_MAIN}v")
                    if (
                        cell.get("t") == "str"
                        and cell.find(f"{_MAIN}f") is not None
                        and value is not None
                        and not value.text
                    ):
                        result.add((str(sheet.get("name", "")), str(cell.get("r", ""))))
                    cell.clear()
                    parent = cell.getparent()
                    if parent is not None:
                        while cell.getprevious() is not None:
                            del parent[0]
    return result
