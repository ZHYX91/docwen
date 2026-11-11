"""Apply authored drawings, notes and links to an existing worksheet package."""

from __future__ import annotations

import re
from copy import deepcopy
from typing import Any

from lxml import etree
from openpyxl.utils.cell import get_column_letter

from .package import PackageEditor, parse_xml, rels_part, resolve_target, serialize_xml

_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_P = "http://schemas.openxmlformats.org/package/2006/relationships"
_D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_V = "urn:schemas-microsoft-com:vml"
_X = "urn:schemas-microsoft-com:office:excel"
_THREADS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
_DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
_COMMENT_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
_VML_CT = "application/vnd.openxmlformats-officedocument.vmlDrawing"

# CT_Worksheet order: object carriers must precede extLst.
_SHEET_ORDER = (
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
)


def _insert_sheet_child(root: Any, child: Any) -> None:
    order = {name: index for index, name in enumerate(_SHEET_ORDER)}
    rank = order[etree.QName(child).localname]
    for index, current in enumerate(root):
        if not isinstance(current.tag, str):
            continue
        if order.get(etree.QName(current).localname, len(order)) > rank:
            root.insert(index, child)
            return
    root.append(child)


def _relationships(parts: dict[str, bytes], owner: str) -> dict[str, Any]:
    payload = parts.get(rels_part(owner))
    if payload is None:
        return {}
    return {item.get("Id", ""): item for item in parse_xml(payload) if item.tag == f"{{{_P}}}Relationship"}


def _related_part(parts: dict[str, bytes], owner: str, kind: str) -> str | None:
    matches = [item for item in _relationships(parts, owner).values() if item.get("Type", "").endswith("/" + kind)]
    if len(matches) > 1:
        raise ValueError(f"Ambiguous XLSX {kind} relationship")
    if not matches:
        return None
    relation = matches[0]
    if relation.get("TargetMode") == "External":
        raise ValueError(f"XLSX {kind} relationship must be internal")
    target = resolve_target(owner, relation.get("Target", ""))
    if target not in parts:
        raise ValueError(f"Missing XLSX {kind} part")
    return target


def _carrier_part(parts: dict[str, bytes], owner: str, root: Any, carrier: str) -> str | None:
    node = root.find(f"{{{_S}}}{carrier}")
    if node is None:
        return None
    relation = _relationships(parts, owner).get(node.get(f"{{{_R}}}id", ""))
    if relation is None or relation.get("TargetMode") == "External":
        raise ValueError(f"Invalid XLSX {carrier} relationship")
    target = resolve_target(owner, relation.get("Target", ""))
    if target not in parts:
        raise ValueError(f"Missing XLSX {carrier} part")
    return target


def _copy_relationship_attributes(
    editor: PackageEditor,
    node: Any,
    destination_owner: str,
    donor_parts: dict[str, bytes],
    donor_owner: str,
    imported: dict[str, str],
    relationships: dict[str, str],
) -> None:
    donor_relationships = _relationships(donor_parts, donor_owner)
    for element in node.iter():
        for attribute, rid in list(element.attrib.items()):
            if etree.QName(attribute).namespace != _R:
                continue
            if rid not in relationships:
                relation = donor_relationships.get(rid)
                if relation is None:
                    raise ValueError("New XLSX object refers to a missing relationship")
                mode = relation.get("TargetMode")
                target = relation.get("Target", "")
                if mode != "External":
                    target = resolve_target(donor_owner, target)
                    if target not in imported:
                        imported[target] = editor.import_part(donor_parts, target)
                    target = imported[target]
                relationships[rid] = editor.add_relationship(destination_owner, relation.get("Type", ""), target, mode)
            element.set(attribute, relationships[rid])


def _merge_images(
    editor: PackageEditor,
    output_part: str,
    root: Any,
    donor_parts: dict[str, bytes],
    donor_part: str,
    donor_root: Any,
    worksheet: Any,
    baseline_images_count: int,
) -> bool:
    image_count = len(worksheet._images)
    if image_count < baseline_images_count:
        raise ValueError("Template image removal must be an explicit package operation")
    added = image_count - baseline_images_count
    if not added:
        return False
    donor_drawing_part = _carrier_part(donor_parts, donor_part, donor_root, "drawing")
    if donor_drawing_part is None:
        raise ValueError("Authored XLSX images have no serialized drawing")
    donor_drawing = parse_xml(donor_parts[donor_drawing_part])
    pictures = [node for node in donor_drawing if node.find(f"{{{_D}}}pic") is not None]
    if len(pictures) < image_count:
        raise ValueError("Authored XLSX image anchors were omitted")
    selected = pictures[-added:]
    drawing_part = _carrier_part(editor.parts, output_part, root, "drawing")
    carrier_added = drawing_part is None
    if drawing_part is None:
        drawing_part = editor.allocate(donor_drawing_part)
        drawing = etree.Element(f"{{{_D}}}wsDr", nsmap=donor_drawing.nsmap)
        editor.parts[drawing_part] = serialize_xml(drawing)
        editor.ensure_content_type(drawing_part, _DRAWING_CT)
        rid = editor.add_relationship(output_part, _R + "/drawing", drawing_part)
        carrier = etree.Element(f"{{{_S}}}drawing")
        carrier.set(f"{{{_R}}}id", rid)
        _insert_sheet_child(root, carrier)
    else:
        drawing = parse_xml(editor.parts[drawing_part])
    used_ids = {element.get("id") for element in drawing.iter(f"{{{_D}}}cNvPr")}
    next_id = max((int(value) for value in used_ids if value and value.isdigit()), default=0) + 1
    imported: dict[str, str] = {}
    relationships: dict[str, str] = {}
    for source in selected:
        anchor = deepcopy(source)
        _copy_relationship_attributes(
            editor, anchor, drawing_part, donor_parts, donor_drawing_part, imported, relationships
        )
        for properties in anchor.iter(f"{{{_D}}}cNvPr"):
            while str(next_id) in used_ids:
                next_id += 1
            properties.set("id", str(next_id))
            used_ids.add(str(next_id))
            next_id += 1
        drawing.append(anchor)
    editor.parts[drawing_part] = serialize_xml(drawing)
    return carrier_added


def _comment_map(root: Any) -> dict[str, Any]:
    return {node.get("ref", ""): node for node in root.findall(f"{{{_S}}}commentList/{{{_S}}}comment")}


def _reject_threaded_comment_conflicts(parts: dict[str, bytes], owner: str, changes: set[str]) -> None:
    if not changes:
        return
    threaded_part = _related_part(parts, owner, "threadedComment")
    if threaded_part is None:
        return
    threaded = parse_xml(parts[threaded_part])
    occupied = {
        node.get("ref", "").replace("$", "").upper()
        for node in threaded.iter(f"{{{_THREADS}}}threadedComment")
        if node.get("ref")
    }
    conflicts = {reference for reference in changes if reference.replace("$", "").upper() in occupied}
    if conflicts:
        raise ValueError(
            "Cannot change a legacy XLSX note where the template has a threaded comment: "
            + ", ".join(sorted(conflicts))
        )


def _note_shape_ref(shape: Any) -> str | None:
    client = shape.find(f"{{{_X}}}ClientData")
    if client is None or client.get("ObjectType") != "Note":
        return None
    row = client.findtext(f"{{{_X}}}Row")
    column = client.findtext(f"{{{_X}}}Column")
    if row is None or column is None:
        return None
    try:
        return f"{get_column_letter(int(column) + 1)}{int(row) + 1}"
    except (ValueError, TypeError):
        return None


def _merge_note_shapes(
    editor: PackageEditor,
    output_part: str,
    root: Any,
    donor_parts: dict[str, bytes],
    donor_part: str,
    donor_root: Any,
    added: set[str],
    removed: set[str],
) -> None:
    if not added and not removed:
        return
    donor_vml_part = _carrier_part(donor_parts, donor_part, donor_root, "legacyDrawing")
    donor_vml = parse_xml(donor_parts[donor_vml_part]) if donor_vml_part else None
    donor_shapes = (
        {_note_shape_ref(node): node for node in donor_vml.iter(f"{{{_V}}}shape")} if donor_vml is not None else {}
    )
    vml_part = _carrier_part(editor.parts, output_part, root, "legacyDrawing")
    if vml_part is None:
        if not added:
            return
        if donor_vml_part is None or donor_vml is None:
            raise ValueError("New XLSX notes have no serialized VML shapes")
        vml_part = editor.allocate(donor_vml_part)
        vml = etree.Element(donor_vml.tag, nsmap=donor_vml.nsmap)
        editor.parts[vml_part] = serialize_xml(vml)
        editor.ensure_content_type(vml_part, _VML_CT)
        rid = editor.add_relationship(output_part, _R + "/vmlDrawing", vml_part)
        carrier = etree.Element(f"{{{_S}}}legacyDrawing")
        carrier.set(f"{{{_R}}}id", rid)
        _insert_sheet_child(root, carrier)
    else:
        vml = parse_xml(editor.parts[vml_part])
    for shape in list(vml.iter(f"{{{_V}}}shape")):
        if _note_shape_ref(shape) in removed:
            parent = shape.getparent()
            if parent is None:
                raise ValueError("A VML note shape has no container")
            parent.remove(shape)
    existing_notes = {_note_shape_ref(node) for node in vml.iter(f"{{{_V}}}shape")}
    used_ids = {identifier for node in vml.iter() if isinstance(identifier := node.get("id"), str)}
    next_shape = (
        max(
            (int(match.group(1)) for item in used_ids if (match := re.fullmatch(r"_x0000_s(\d+)", item))),
            default=1024,
        )
        + 1
    )
    for reference in sorted(added - existing_notes):
        source_shape = donor_shapes.get(reference)
        if source_shape is None or donor_vml is None:
            raise ValueError("New XLSX note has no matching VML shape")
        shape = deepcopy(source_shape)
        type_reference = shape.get("type", "").removeprefix("#")
        donor_type = next(
            (node for node in donor_vml.iter(f"{{{_V}}}shapetype") if node.get("id") == type_reference), None
        )
        existing_type = next(
            (node for node in vml.iter(f"{{{_V}}}shapetype") if node.get("id") == type_reference), None
        )
        if donor_type is not None and existing_type is None:
            vml.insert(0, deepcopy(donor_type))
            used_ids.add(type_reference)
        elif (
            donor_type is not None
            and existing_type is not None
            and etree.tostring(donor_type) != etree.tostring(existing_type)
        ):
            new_type = deepcopy(donor_type)
            suffix = 1
            while f"docwen_note_type_{suffix}" in used_ids:
                suffix += 1
            type_reference = f"docwen_note_type_{suffix}"
            new_type.set("id", type_reference)
            used_ids.add(type_reference)
            vml.insert(0, new_type)
        if type_reference:
            shape.set("type", "#" + type_reference)
        while f"_x0000_s{next_shape}" in used_ids:
            next_shape += 1
        shape_id = f"_x0000_s{next_shape}"
        shape.set("id", shape_id)
        used_ids.add(shape_id)
        next_shape += 1
        vml.append(shape)
    editor.parts[vml_part] = serialize_xml(vml)


def _merge_comments(
    editor: PackageEditor,
    output_part: str,
    root: Any,
    donor_parts: dict[str, bytes],
    donor_part: str,
    donor_root: Any,
    changes: set[str],
) -> bool:
    if not changes:
        return False
    donor_comment_part = _related_part(donor_parts, donor_part, "comments")
    donor_comments = parse_xml(donor_parts[donor_comment_part]) if donor_comment_part else None
    donor_map = _comment_map(donor_comments) if donor_comments is not None else {}
    source_part = _related_part(editor.parts, output_part, "comments")
    if source_part is None:
        if not changes.intersection(donor_map):
            return False
        source_part = editor.allocate(donor_comment_part or "xl/comments/comment1.xml")
        comments = parse_xml(f'<comments xmlns="{_S}"/>'.encode())
        etree.SubElement(comments, f"{{{_S}}}authors")
        etree.SubElement(comments, f"{{{_S}}}commentList")
        editor.parts[source_part] = serialize_xml(comments)
        editor.ensure_content_type(source_part, _COMMENT_CT)
        editor.add_relationship(output_part, _R + "/comments", source_part)
    else:
        comments = parse_xml(editor.parts[source_part])
    source_map = _comment_map(comments)
    authors = comments.find(f"{{{_S}}}authors")
    comment_list = comments.find(f"{{{_S}}}commentList")
    if authors is None or comment_list is None:
        raise ValueError("XLSX comments are missing their authors or comment list")
    source_authors = [item.text or "" for item in authors]
    donor_authors = (
        [item.text or "" for item in donor_comments.findall(f"{{{_S}}}authors/{{{_S}}}author")]
        if donor_comments is not None
        else []
    )
    added: set[str] = set()
    removed: set[str] = set()
    for reference in sorted(changes):
        donor = donor_map.get(reference)
        original = source_map.get(reference)
        if donor is None:
            if original is not None:
                comment_list.remove(original)
                removed.add(reference)
            continue
        author_id = int(donor.get("authorId", "0"))
        if author_id >= len(donor_authors):
            raise ValueError("Authored XLSX note has an invalid author")
        author = donor_authors[author_id]
        if author not in source_authors:
            etree.SubElement(authors, f"{{{_S}}}author").text = author
            source_authors.append(author)
        if original is None:
            original = deepcopy(donor)
            comment_list.append(original)
            added.add(reference)
        else:
            old_text = original.find(f"{{{_S}}}text")
            new_text = donor.find(f"{{{_S}}}text")
            if old_text is not None:
                original.remove(old_text)
            if new_text is not None:
                original.insert(0, deepcopy(new_text))
        original.set("authorId", str(source_authors.index(author)))
    had_vml_carrier = root.find(f"{{{_S}}}legacyDrawing") is not None
    editor.parts[source_part] = serialize_xml(comments)
    _merge_note_shapes(editor, output_part, root, donor_parts, donor_part, donor_root, added, removed)
    return not had_vml_carrier and root.find(f"{{{_S}}}legacyDrawing") is not None


def _merge_hyperlinks(
    editor: PackageEditor,
    output_part: str,
    root: Any,
    donor_parts: dict[str, bytes],
    donor_part: str,
    donor_root: Any,
    changes: set[str],
) -> bool:
    if not changes:
        return False
    source_links = root.find(f"{{{_S}}}hyperlinks")
    donor_links = donor_root.find(f"{{{_S}}}hyperlinks")
    donor_map = {node.get("ref", ""): node for node in donor_links} if donor_links is not None else {}
    if source_links is None:
        if not changes.intersection(donor_map):
            return False
        source_links = etree.Element(f"{{{_S}}}hyperlinks")
        _insert_sheet_child(root, source_links)
    source_map = {node.get("ref", ""): node for node in source_links}
    imported: dict[str, str] = {}
    relationships: dict[str, str] = {}
    for reference in sorted(changes):
        old = source_map.get(reference)
        new = donor_map.get(reference)
        if old is not None:
            source_links.remove(old)
        if new is not None:
            link = deepcopy(new)
            _copy_relationship_attributes(editor, link, output_part, donor_parts, donor_part, imported, relationships)
            source_links.append(link)
    if not len(source_links):
        root.remove(source_links)
    return True


def merge_sheet_objects(
    editor: PackageEditor,
    output_part: str,
    root: Any,
    donor_parts: dict[str, bytes],
    donor_part: str,
    donor_root: Any,
    worksheet: Any,
    baseline_images_count: int,
    comment_changes: set[str],
    hyperlink_changes: set[str],
    *,
    is_new: bool = False,
) -> bool:
    """Merge only objects deliberately changed since the loaded snapshot.

    New sheets are imported as complete relationship graphs by the caller;
    existing sheets (including raw-template clones) use the narrow operations
    here so unsupported original drawing/VML/extension content survives.
    """
    if is_new:
        return False
    _reject_threaded_comment_conflicts(editor.parts, output_part, comment_changes)
    changed = _merge_images(
        editor, output_part, root, donor_parts, donor_part, donor_root, worksheet, baseline_images_count
    )
    changed = (
        _merge_comments(editor, output_part, root, donor_parts, donor_part, donor_root, comment_changes) or changed
    )
    return (
        _merge_hyperlinks(editor, output_part, root, donor_parts, donor_part, donor_root, hyperlink_changes) or changed
    )
