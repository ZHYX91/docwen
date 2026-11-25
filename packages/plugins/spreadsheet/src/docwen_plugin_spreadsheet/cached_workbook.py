"""Inspect stored worksheet formula cells without materializing sparse grids."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path
from typing import TYPE_CHECKING
from zipfile import ZipFile

from docwen_core.spreadsheet_package import ordinary_worksheet_parts, preflight_worksheet_allocations

if TYPE_CHECKING:
    from openpyxl import Workbook

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def load_cached_workbook(path: str, *, cancel_check: Callable[[], None], limit: int) -> Workbook:
    """Check allocation budgets before opening the cached values of admitted XLSX content."""
    import openpyxl

    preflight_worksheet_allocations(path, cancel_check=cancel_check, limit=limit)
    # File admission is content-based. Passing the stream preserves that decision
    # even when the source name has a misleading extension.
    with Path(path).open("rb") as stream:
        return openpyxl.load_workbook(stream, data_only=True)


def find_unavailable_formula_caches(
    values_workbook: Workbook,
    *,
    input_path: str,
    sheet_names: list[str] | None = None,
    location_limit: int = 20,
    cancel_check: Callable[[], None] | None = None,
) -> tuple[int, list[str]]:
    """Find formula cells whose cached scalar value is unavailable to openpyxl."""

    count = 0
    locations: list[str] = []

    def visit(sheet_name: str, coordinate: str, known_empty: bool) -> None:
        nonlocal count
        if known_empty or values_workbook[sheet_name][coordinate].value is not None:
            return
        count += 1
        if len(locations) < location_limit:
            locations.append(f"{sheet_name}!{coordinate}")

    visit_formula_caches(input_path, visit, sheet_names=sheet_names, cancel_check=cancel_check)
    return count, locations


def formula_cache_warning_message(count: int, locations: list[str]) -> str:
    shown = ", ".join(locations)
    omitted = count - len(locations)
    suffix = f"; {omitted} more not listed" if omitted > 0 else ""
    return (
        f"{count} formula cell(s) have no usable cached value in this workbook and were exported as empty text. "
        f"Locations: {shown}{suffix}. Recalculate and save the workbook in a spreadsheet application when values "
        "are required. This warning does not establish whether any existing cached values are current."
    )


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
    from lxml import etree
    from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

    selected = None if sheet_names is None else set(sheet_names)
    with ZipFile(path) as package:
        for name, part in ordinary_worksheet_parts(package):
            if cancel_check:
                cancel_check()
            if selected is not None and name not in selected:
                continue
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
