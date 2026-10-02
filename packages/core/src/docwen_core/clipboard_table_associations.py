"""Reversible HTML table-header associations carried in DOCX custom XML.

The carrier is private to managed structured clipboard documents. It reuses the
repository's canonical custom-XML package infrastructure but does not change
resolved-v4, numbering, or public Machine Protocol schemas.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import lxml.etree as etree

from docwen_core.models.clipboard_document import (
    ClipboardDocument,
    ClipboardTable,
    clipboard_table_header_shape,
    iter_clipboard_blocks,
)

CLIPBOARD_TABLE_ASSOCIATION_MAP_NAMESPACE = "urn:docwen:clipboard-table-associations:v1"
_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
_ITEM_RE = re.compile(r"customXml/item([1-9][0-9]*)\.xml$")
_SCOPES = frozenset({"", "row", "col", "rowgroup", "colgroup"})


class ClipboardTableAssociationError(ValueError):
    """The reversible clipboard-table association carrier is invalid."""


@dataclass(frozen=True, slots=True)
class ClipboardCellAssociation:
    row: int
    column: int
    row_span: int
    column_span: int
    header: bool
    scope: str
    html_id: str
    headers: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class ClipboardTableAssociation:
    index: int
    header_rows: int
    header_columns: int
    cells: tuple[ClipboardCellAssociation, ...]


def _xml_attr(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace('"', "&quot;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\r", "&#13;")
        .replace("\n", "&#10;")
        .replace("\t", "&#9;")
    )


def build_clipboard_table_associations(document: ClipboardDocument) -> tuple[ClipboardTableAssociation, ...]:
    """Build table association records in recursive document preorder."""

    tables = [block for block in iter_clipboard_blocks(document) if isinstance(block, ClipboardTable)]
    output: list[ClipboardTableAssociation] = []
    for index, table in enumerate(tables):
        header_rows, header_columns = clipboard_table_header_shape(table)
        cells = tuple(
            ClipboardCellAssociation(
                row=cell.row,
                column=cell.column,
                row_span=cell.row_span,
                column_span=cell.column_span,
                header=cell.header,
                scope=cell.scope,
                html_id=cell.html_id,
                headers=cell.headers,
            )
            for cell in sorted(table.cells, key=lambda item: (item.row, item.column))
        )
        output.append(ClipboardTableAssociation(index, header_rows, header_columns, cells))
    return tuple(output)


def clipboard_table_association_map_xml(records: tuple[ClipboardTableAssociation, ...]) -> bytes:
    """Serialize the closed association map canonically."""

    table_xml: list[str] = []
    for table in records:
        cells: list[str] = []
        for cell in table.cells:
            header_refs = "".join(f'<header ref="{_xml_attr(item)}"/>' for item in cell.headers)
            cells.append(
                f'<cell row="{cell.row}" column="{cell.column}" row_span="{cell.row_span}" '
                f'column_span="{cell.column_span}" header="{"1" if cell.header else "0"}" '
                f'scope="{_xml_attr(cell.scope)}" html_id="{_xml_attr(cell.html_id)}">'
                f"{header_refs}</cell>"
            )
        table_xml.append(
            f'<table index="{table.index}" header_rows="{table.header_rows}" '
            f'header_columns="{table.header_columns}">{"".join(cells)}</table>'
        )
    root = (
        f'<clipboardTableAssociations xmlns="{CLIPBOARD_TABLE_ASSOCIATION_MAP_NAMESPACE}" version="1">'
        f"{''.join(table_xml)}</clipboardTableAssociations>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def _require_int(raw: str | None, *, minimum: int, field: str) -> int:
    if raw is None or re.fullmatch(r"0|[1-9][0-9]*", raw) is None:
        raise ClipboardTableAssociationError(f"{field} is not a canonical non-negative integer")
    value = int(raw)
    if value < minimum:
        raise ClipboardTableAssociationError(f"{field} is out of range")
    return value


def parse_clipboard_table_association_map(root: Any) -> tuple[ClipboardTableAssociation, ...]:
    """Parse one canonical clipboard table-association map."""

    namespace = f"{{{CLIPBOARD_TABLE_ASSOCIATION_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}clipboardTableAssociations"
        or tuple(root.attrib.items()) != (("version", "1"),)
        or root.text is not None
        or root.tail is not None
    ):
        raise ClipboardTableAssociationError("clipboard table-association root is invalid")

    tables: list[ClipboardTableAssociation] = []
    for expected_index, table_node in enumerate(root):
        if (
            table_node.tag != f"{namespace}table"
            or tuple(table_node.attrib) != ("index", "header_rows", "header_columns")
            or table_node.text is not None
            or table_node.tail is not None
        ):
            raise ClipboardTableAssociationError("clipboard table-association table entry is invalid")
        index = _require_int(table_node.get("index"), minimum=0, field="table.index")
        if index != expected_index:
            raise ClipboardTableAssociationError("clipboard table indices are not contiguous")
        header_rows = _require_int(table_node.get("header_rows"), minimum=0, field="table.header_rows")
        header_columns = _require_int(table_node.get("header_columns"), minimum=0, field="table.header_columns")
        cells: list[ClipboardCellAssociation] = []
        previous_position: tuple[int, int] | None = None
        for cell_node in table_node:
            if (
                cell_node.tag != f"{namespace}cell"
                or tuple(cell_node.attrib) != ("row", "column", "row_span", "column_span", "header", "scope", "html_id")
                or cell_node.text is not None
                or cell_node.tail is not None
            ):
                raise ClipboardTableAssociationError("clipboard table-association cell entry is invalid")
            row = _require_int(cell_node.get("row"), minimum=0, field="cell.row")
            column = _require_int(cell_node.get("column"), minimum=0, field="cell.column")
            row_span = _require_int(cell_node.get("row_span"), minimum=1, field="cell.row_span")
            column_span = _require_int(cell_node.get("column_span"), minimum=1, field="cell.column_span")
            position = (row, column)
            if previous_position is not None and position <= previous_position:
                raise ClipboardTableAssociationError("clipboard table cells are not canonically ordered")
            previous_position = position
            header_raw = cell_node.get("header")
            if header_raw not in {"0", "1"}:
                raise ClipboardTableAssociationError("clipboard table header flag is invalid")
            scope = cell_node.get("scope") or ""
            if scope not in _SCOPES:
                raise ClipboardTableAssociationError("clipboard table scope is invalid")
            html_id = cell_node.get("html_id") or ""
            headers: list[str] = []
            for header_node in cell_node:
                if (
                    header_node.tag != f"{namespace}header"
                    or tuple(header_node.attrib) != ("ref",)
                    or not header_node.get("ref")
                    or header_node.text is not None
                    or header_node.tail is not None
                    or len(header_node) != 0
                ):
                    raise ClipboardTableAssociationError("clipboard table headers reference is invalid")
                headers.append(str(header_node.get("ref")))
            cells.append(
                ClipboardCellAssociation(
                    row=row,
                    column=column,
                    row_span=row_span,
                    column_span=column_span,
                    header=header_raw == "1",
                    scope=scope,
                    html_id=html_id,
                    headers=tuple(headers),
                )
            )
        tables.append(ClipboardTableAssociation(index, header_rows, header_columns, tuple(cells)))
    return tuple(tables)


def inject_clipboard_table_associations(path: Path, document: ClipboardDocument) -> None:
    """Inject one canonical custom-XML map while preserving unrelated semantic maps."""

    records = build_clipboard_table_associations(document)
    if not records:
        return
    with ZipFile(path) as package:
        matches = []
        parser = etree.XMLParser(resolve_entities=False, no_network=True)
        for name in package.namelist():
            if _ITEM_RE.fullmatch(name) is None:
                continue
            try:
                root = etree.fromstring(package.read(name), parser)
            except etree.XMLSyntaxError:
                continue
            if etree.QName(root).namespace == CLIPBOARD_TABLE_ASSOCIATION_MAP_NAMESPACE:
                matches.append(name)
        if matches:
            raise ClipboardTableAssociationError("DOCX package already contains clipboard table associations")

    from docwen_core._docx_semantics_v3_package import inject_custom_xml_parts

    inject_custom_xml_parts(
        path,
        [(CLIPBOARD_TABLE_ASSOCIATION_MAP_NAMESPACE, clipboard_table_association_map_xml(records))],
        allow_existing_owned=True,
    )


def read_clipboard_table_associations(path: Path) -> tuple[ClipboardTableAssociation, ...]:
    """Read and fully verify the reversible clipboard table-association map."""

    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    with ZipFile(path) as package:
        matches: list[tuple[int, Any]] = []
        for name in package.namelist():
            match = _ITEM_RE.fullmatch(name)
            if match is None:
                continue
            try:
                root = etree.fromstring(package.read(name), parser)
            except etree.XMLSyntaxError:
                continue
            if etree.QName(root).namespace == CLIPBOARD_TABLE_ASSOCIATION_MAP_NAMESPACE:
                matches.append((int(match.group(1)), root))
        if not matches:
            return ()
        if len(matches) != 1:
            raise ClipboardTableAssociationError("DOCX package contains duplicate clipboard table associations")
        item_number, root = matches[0]

        from docwen_core._docx_semantics_v3_package import verify_custom_xml_support

        verify_custom_xml_support(package, item_number, CLIPBOARD_TABLE_ASSOCIATION_MAP_NAMESPACE)
        return parse_clipboard_table_association_map(root)


__all__ = [
    "CLIPBOARD_TABLE_ASSOCIATION_MAP_NAMESPACE",
    "ClipboardCellAssociation",
    "ClipboardTableAssociation",
    "ClipboardTableAssociationError",
    "build_clipboard_table_associations",
    "clipboard_table_association_map_xml",
    "inject_clipboard_table_associations",
    "parse_clipboard_table_association_map",
    "read_clipboard_table_associations",
]
