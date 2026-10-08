"""Word-save-stable table roles, bound to physical tables and merge geometry.

Conditional style flags are presentation hints: Word rewrites them on save.
This private custom XML carrier preserves the authored role counts while native
text, merge geometry, repeat policy and explicit table-look edits remain live.
"""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from typing import Any, cast
from zipfile import ZipFile

import lxml.etree as etree

from docwen_core.docx_bookmarks import BOOKMARK_ID_MAX, build_docx_bookmark_inventory, prove_bookmark_name

TABLE_ROLES_NAMESPACE = "urn:docwen:table-roles:v1"
_NAMESPACE = TABLE_ROLES_NAMESPACE
_PREFIX = "_DWT_"
_ITEM = re.compile(r"customXml/item([1-9][0-9]*)\.xml$")
_FIELDS = {"bookmark", "rows", "columns", "shape"}


def table_role_bookmark_nodes(element: Any) -> frozenset[Any]:
    """Recognize only adjacent, zero-width auxiliary pairs in first cells.

    This narrow exception lets ordinary table anchors retain their no-fields
    rule. Package recovery separately proves global uniqueness and map binding.
    """
    from docx.oxml.ns import qn

    allowed: set[Any] = set()
    for table in element.iter(qn("w:tbl")):
        anchor = table.find(f"{qn('w:tr')}/{qn('w:tc')}/{qn('w:p')}")
        if anchor is None:
            continue
        children = list(anchor)
        for index, start in enumerate(children[:-1]):
            if (
                start.tag != qn("w:bookmarkStart")
                or re.fullmatch(r"_DWT_[0-9a-f]{32}", start.get(qn("w:name"), "")) is None
            ):
                continue
            end = children[index + 1]
            raw_id = start.get(qn("w:id"), "")
            if (
                end.tag == qn("w:bookmarkEnd")
                and end.get(qn("w:id")) == raw_id
                and re.fullmatch(r"0|[1-9][0-9]{0,9}", raw_id)
                and int(raw_id) <= BOOKMARK_ID_MAX
            ):
                allowed.update((start, end))
    return frozenset(allowed)


def prove_table_role_bookmarks(document: Any, root: Any | None) -> frozenset[Any]:
    """Authorize auxiliary pairs only against a validated closed role map.

    Deleted declarations are handled as stale by recovery. Existing auxiliary
    bookmarks must all be declared, globally unique, and in the exact slot.
    A missing map grants no exception to ordinary-anchor bookmark proof.
    """
    if root is None:
        return frozenset()
    canonical_table_roles_xml(root)
    declared = {record.get("bookmark").casefold() for record in root}
    inventory = build_docx_bookmark_inventory(document)
    structural_nodes = table_role_bookmark_nodes(document.element)
    proven: set[Any] = set()
    for start in inventory.starts:
        name = start.name or ""
        if not name.casefold().startswith(_PREFIX.casefold()):
            continue
        if name.casefold() not in declared:
            raise ValueError("Undeclared table role bookmark")
        proof = prove_bookmark_name(inventory, name)
        if (
            not proof.valid
            or proof.start is None
            or proof.end is None
            or proof.start.element not in structural_nodes
            or proof.end.element not in structural_nodes
        ):
            raise ValueError("Invalid table role bookmark binding or position")
        proven.update((proof.start.element, proof.end.element))
    return frozenset(proven)


def canonical_table_roles_xml(root: Any) -> bytes:
    """Validate the closed map and produce the package identity bytes."""
    if root.tag != f"{{{_NAMESPACE}}}tableRoles" or dict(root.attrib) != {"version": "1"} or root.text or root.tail:
        raise ValueError("Invalid table role map")
    canonical = etree.Element(f"{{{_NAMESPACE}}}tableRoles", nsmap=cast(Any, {None: _NAMESPACE}), version="1")
    seen: set[str] = set()
    for record in root:
        if (
            record.tag != f"{{{_NAMESPACE}}}table"
            or set(record.attrib) != _FIELDS
            or len(record)
            or record.text
            or record.tail
        ):
            raise ValueError("Invalid table role record")
        name = record.get("bookmark", "")
        if re.fullmatch(r"_DWT_[0-9a-f]{32}", name) is None or name in seen:
            raise ValueError("Invalid or duplicate table role binding")
        seen.add(name)
        if any(re.fullmatch(r"0|[1-9][0-9]{0,8}", record.get(key, "")) is None for key in ("rows", "columns")):
            raise ValueError("Invalid table role counts")
        if re.fullmatch(r"[0-9a-f]{64}", record.get("shape", "")) is None:
            raise ValueError("Invalid table role geometry digest")
        etree.SubElement(
            canonical, record.tag, **{key: record.get(key) for key in ("bookmark", "rows", "columns", "shape")}
        )
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + etree.tostring(canonical) + b"\n"


def _shape(table: Any) -> tuple[str, int, int]:
    from docwen_core.docx_parsing.table_extraction import build_docx_table_semantic_grid

    grid = build_docx_table_semantic_grid(table, cell_text_resolver=lambda _cell, _row, _col: "")
    geometry = [[(cell.anchor_row, cell.anchor_col, cell.rowspan, cell.colspan) for cell in row] for row in grid]
    digest = hashlib.sha256(json.dumps(geometry, separators=(",", ":")).encode()).hexdigest()
    return digest, len(grid), len(grid[0])


def prepare_table_roles(document: Any, *, tables: tuple[Any, ...] | None = None) -> bytes | None:
    """Bind final rendered tables before saving; do not depend on table order."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    from docwen_core.docx_parsing.document_semantics import extract_semantic_table_metadata

    inventory = build_docx_bookmark_inventory(document)
    used_ids = set(inventory.used_id_keys)
    used_names = set(inventory.used_name_keys)
    root = etree.Element(f"{{{_NAMESPACE}}}tableRoles", nsmap=cast(Any, {None: _NAMESPACE}), version="1")
    next_id = 0
    for table in tables if tables is not None else document.element.body.iter(qn("w:tbl")):
        # Only tables with explicit role evidence need an authored carrier.
        if not list(table.iter(qn("w:cnfStyle"))):
            continue
        roles = extract_semantic_table_metadata(table, default_first_row=False)
        # Native first-row/explicit-zero look is sufficient for ordinary tables.
        if roles.header_rows <= 1 and not roles.header_columns:
            continue
        digest, row_count, column_count = _shape(table)
        if roles.header_rows > row_count or roles.header_columns > column_count:
            raise ValueError("Table roles exceed the physical table geometry")
        anchor = table.find(f"{qn('w:tr')}/{qn('w:tc')}/{qn('w:p')}")
        if anchor is None:
            raise ValueError("Table role binding requires a first-cell paragraph")
        existing = [
            node for node in anchor.findall(qn("w:bookmarkStart")) if node.get(qn("w:name"), "").startswith(_PREFIX)
        ]
        if existing:
            if len(existing) != 1:
                raise ValueError("Multiple table role bookmarks in export table")
            name = existing[0].get(qn("w:name"))
            if not prove_bookmark_name(inventory, name, scope_element=table).valid:
                raise ValueError("Invalid prebound table role bookmark")
        else:
            while ("numeric", next_id) in used_ids:
                next_id += 1
            salt = next_id
            while True:
                identity = f"{digest}:{roles.header_rows}:{roles.header_columns}:{salt}"
                name = _PREFIX + hashlib.sha256(identity.encode()).hexdigest()[:32]
                if name.casefold() not in used_names:
                    break
                salt += 1
            used_names.add(name.casefold())
            if next_id > BOOKMARK_ID_MAX:
                raise ValueError("Table role bookmark identifiers exhausted")
            used_ids.add(("numeric", next_id))
            start = OxmlElement("w:bookmarkStart")
            start.set(qn("w:id"), str(next_id))
            start.set(qn("w:name"), name)
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), str(next_id))
            offset = 1 if len(anchor) and anchor[0].tag == qn("w:pPr") else 0
            anchor.insert(offset, start)
            anchor.insert(offset + 1, end)
        etree.SubElement(
            root,
            f"{{{_NAMESPACE}}}table",
            bookmark=name,
            rows=str(roles.header_rows),
            columns=str(roles.header_columns),
            shape=digest,
        )
    return canonical_table_roles_xml(root) if len(root) else None


def inject_table_roles(path: Path, payload: bytes | None) -> None:
    if payload is None:
        return
    from docwen_core._docx_semantics_v3_package import inject_custom_xml_parts

    inject_custom_xml_parts(path, [(_NAMESPACE, payload)], allow_existing_owned=True)


def recover_table_roles(
    path: Path, document: Any, *, role_overrides: dict[Any, tuple[int, int]] | None = None
) -> tuple[str, ...]:
    """Validate bindings before restoring volatile style hints in memory only.

    Edited/deleted tables use native roles with an explicit warning. Ambiguous
    bindings fail closed; stale records never transfer by table order or text.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    from docwen_core._docx_semantics_v3_package import verify_custom_xml_support

    if role_overrides is not None:
        role_overrides.clear()
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=True)
    maps: list[Any] = []
    with ZipFile(path) as package:
        if len(package.namelist()) != len(set(package.namelist())):
            raise ValueError("DOCX contains duplicate ZIP members")
        for name in package.namelist():
            match = _ITEM.fullmatch(name)
            if match is None:
                continue
            data = package.read(name)
            try:
                root = etree.fromstring(data, parser)
            except etree.XMLSyntaxError as exc:
                if _NAMESPACE.encode() in data:
                    raise ValueError("Malformed table role map") from exc
                continue
            if etree.QName(root).namespace == _NAMESPACE:
                verify_custom_xml_support(package, int(match.group(1)), _NAMESPACE)
                maps.append(root)
    if not maps:
        if any((item.name or "").startswith(_PREFIX) for item in build_docx_bookmark_inventory(document).starts):
            return ("Authored table role metadata is missing; native table roles were used.",)
        return ()
    if len(maps) != 1:
        raise ValueError("Duplicate table role maps")
    root = maps[0]
    prove_table_role_bookmarks(document, root)
    if root.tag != f"{{{_NAMESPACE}}}tableRoles" or dict(root.attrib) != {"version": "1"} or root.text:
        raise ValueError("Invalid table role map")
    inventory = build_docx_bookmark_inventory(document)
    warnings: list[str] = []
    seen: set[str] = set()
    validated: list[tuple[Any, int, int]] = []
    for record in root:
        if record.tag != f"{{{_NAMESPACE}}}table" or set(record.attrib) != _FIELDS or len(record) or record.text:
            raise ValueError("Invalid table role record")
        name = record.get("bookmark", "")
        if re.fullmatch(r"_DWT_[0-9a-f]{32}", name) is None or name in seen:
            raise ValueError("Invalid or duplicate table role binding")
        seen.add(name)
        if not inventory.starts_named(name):
            warnings.append("An authored table role binding was deleted; no role record was applied for that binding.")
            continue
        proof = prove_bookmark_name(inventory, name)
        if not proof.valid or proof.start is None or proof.end is None:
            raise ValueError("Table role bookmark is missing or ambiguous")
        table = next((node for node in proof.start.element.iterancestors() if node.tag == qn("w:tbl")), None)
        end_table = next((node for node in proof.end.element.iterancestors() if node.tag == qn("w:tbl")), None)
        if table is None or end_table is not table:
            raise ValueError("Table role bookmark crosses physical tables")
        if any(existing is table for existing, _rows, _columns in validated):
            raise ValueError("Multiple role records bind the same table")
        digest, row_count, column_count = _shape(table)
        if record.get("shape") != digest:
            warnings.append(
                "Table geometry changed; stale authored roles were ignored and native table roles were used."
            )
            continue
        raw_rows, raw_columns = record.get("rows", ""), record.get("columns", "")
        if any(re.fullmatch(r"0|[1-9][0-9]{0,8}", value) is None for value in (raw_rows, raw_columns)):
            raise ValueError("Invalid table role counts")
        rows, columns = int(raw_rows), int(raw_columns)
        if rows > row_count or columns > column_count:
            raise ValueError("Table role counts exceed physical geometry")
        validated.append((table, rows, columns))

    for table, header_rows, header_columns in validated:
        look = table.find(f"{qn('w:tblPr')}/{qn('w:tblLook')}")
        # A native explicit disable wins over the authored carrier.
        if look is not None:
            if (look.get(qn("w:firstRow")) or "").casefold() in {"0", "false", "off"}:
                header_rows = 0
            if (look.get(qn("w:firstColumn")) or "").casefold() in {"0", "false", "off"}:
                header_columns = 0
        if role_overrides is not None:
            role_overrides[table] = (header_rows, header_columns)
        for row_index, row in enumerate(table.findall(qn("w:tr"))):
            properties = row.find(qn("w:trPr"))
            if properties is None:
                properties = OxmlElement("w:trPr")
                row.insert(0, properties)
            conditional = properties.find(qn("w:cnfStyle"))
            if conditional is None:
                conditional = OxmlElement("w:cnfStyle")
                properties.insert(0, conditional)
            conditional.set(qn("w:firstRow"), "1" if row_index < header_rows else "0")
            column_index = 0
            for cell in row.findall(qn("w:tc")):
                cell_properties = cell.find(qn("w:tcPr"))
                if cell_properties is None:
                    cell_properties = OxmlElement("w:tcPr")
                    cell.insert(0, cell_properties)
                conditional = cell_properties.find(qn("w:cnfStyle"))
                if conditional is None:
                    conditional = OxmlElement("w:cnfStyle")
                    cell_properties.insert(0, conditional)
                conditional.set(qn("w:firstColumn"), "1" if column_index < header_columns else "0")
                span = cell_properties.find(qn("w:gridSpan"))
                column_index += int(span.get(qn("w:val"), "1")) if span is not None else 1
    return tuple(warnings)
