"""Write XLSX edits against the original template package."""

from __future__ import annotations

import io
import os
import tempfile
from copy import deepcopy
from dataclasses import replace
from pathlib import Path
from typing import Any
from urllib.parse import unquote
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from openpyxl import load_workbook
from openpyxl.styles.stylesheet import write_stylesheet

from docwen_plugin_markdown.xlsx_text import check_workbook_text_capacity

from .state import SheetOrigin, SheetState, TemplateSession
from .styles import StyleMerger, capture_style_state

_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_C = "http://schemas.openxmlformats.org/package/2006/content-types"
_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_SESSION = "_docwen_template_session"


def _parse(payload: bytes) -> Any:
    return etree.fromstring(payload, etree.XMLParser(resolve_entities=False, no_network=True))


def _xml(root: Any) -> bytes:
    return etree.tostring(root, encoding="utf-8", xml_declaration=True)


def _parts(payload: bytes) -> dict[str, bytes]:
    with ZipFile(io.BytesIO(payload)) as package:
        names = package.namelist()
        if len(names) != len(set(names)):
            raise ValueError("XLSX contains duplicate package members")
        return {name: package.read(name) for name in names}


def _target(owner: str, target: str) -> str:
    from .package import resolve_target

    return resolve_target(owner, target)


def _worksheets(parts: dict[str, bytes], *, all_sheets: bool = False) -> dict[str, str]:
    relations = _parse(parts["xl/_rels/workbook.xml.rels"])
    targets = {
        relation.get("Id"): _target("xl/workbook.xml", relation.get("Target", ""))
        for relation in relations
        if (all_sheets or relation.get("Type", "").endswith("/worksheet")) and relation.get("TargetMode") != "External"
    }
    workbook = _parse(parts["xl/workbook.xml"])
    return {
        sheet.get("name"): targets[sheet.get(f"{{{_R}}}id")]
        for sheet in workbook.findall(f"{{{_S}}}sheets/{{{_S}}}sheet")
        if sheet.get(f"{{{_R}}}id") in targets
    }


def _ordered_sheets(workbook: Any) -> tuple[Any, ...]:
    return tuple(workbook[name] for name in workbook.sheetnames)


def _source_sheet(session: TemplateSession, origin: SheetOrigin) -> Any:
    for sheet in session.original_sheets:
        source_origin = session.origins.get(sheet)
        if source_origin is not None and source_origin.part == origin.part:
            return sheet
    raise ValueError("Template clone has no original worksheet owner")


def load_template_workbook(path: str | Path) -> Any:
    """Load the edit model and retain the exact source package independently."""
    source = Path(path)
    payload = source.read_bytes()
    parts = _parts(payload)
    workbook = load_workbook(io.BytesIO(payload))
    sheets = _worksheets(parts)
    session = TemplateSession(
        parts,
        str(source.resolve()),
        {sheet: SheetOrigin(sheets[sheet.title], SheetState.capture(sheet)) for sheet in workbook.worksheets},
        capture_style_state(workbook),
        parts.get("xl/styles.xml") or _xml(write_stylesheet(workbook)),
        _ordered_sheets(workbook),
        {sheet: sheet.title for sheet in _ordered_sheets(workbook)},
    )
    setattr(workbook, _SESSION, session)
    return workbook


def register_template_clone(workbook: Any, target_ws: Any, source_ws: Any) -> None:
    """Associate a copied worksheet with its original package before editing it."""
    session: TemplateSession | None = getattr(workbook, _SESSION, None)
    if session is None:
        return
    if source_ws not in session.origins or target_ws.parent is not workbook:
        raise ValueError("Template worksheet clone must belong to the same workbook")
    origin = session.origins[source_ws]
    # copy_worksheet does not copy drawings. Raw template drawings still belong
    # to the clone, while every drawing actually added to the new model is new.
    baseline = replace(origin.baseline, image_count=len(target_ws._images), title=target_ws.title)
    session.origins[target_ws] = SheetOrigin(origin.part, baseline, clone=True)


def _empty_strings(parts: dict[str, bytes], workbook: Any) -> None:
    sheets = _worksheets(parts)
    for sheet in workbook.worksheets:
        empty = {cell.coordinate for cell in sheet._cells.values() if cell.data_type == "s" and cell.value == ""}
        if not empty:
            continue
        part = sheets[sheet.title]
        root = _parse(parts[part])
        found: set[str] = set()
        for cell in root.findall(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c"):
            if cell.get("r") not in empty:
                continue
            cell.set("t", "inlineStr")
            for child in list(cell):
                if child.tag in {f"{{{_S}}}{name}" for name in ("f", "v", "is")}:
                    cell.remove(child)
            text = etree.SubElement(etree.SubElement(cell, f"{{{_S}}}is"), f"{{{_S}}}t")
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            text.text = ""
            found.add(cell.get("r"))
        if found != empty:
            raise ValueError("An explicitly authored empty XLSX string was not serialized")
        parts[part] = _xml(root)


def _map_sheet_styles(root: Any, styles: StyleMerger) -> None:
    for cell in root.findall(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c"):
        if "s" in cell.attrib:
            cell.set("s", str(styles.map_style(int(cell.get("s")))))
    for row in root.findall(f"{{{_S}}}sheetData/{{{_S}}}row"):
        if "s" in row.attrib:
            row.set("s", str(styles.map_style(int(row.get("s")))))
    for column in root.findall(f"{{{_S}}}cols/{{{_S}}}col"):
        if "style" in column.attrib:
            column.set("style", str(styles.map_style(int(column.get("style")))))


def _defined_names(root: Any, donor_root: Any, session: TemplateSession, workbook: Any) -> bool:
    from .package import relocate_formula_references

    names = root.find(f"{{{_S}}}definedNames")
    changed = False
    current = list(_ordered_sheets(workbook))
    keys: set[tuple[str, str | None]] = set()
    original_scoped = (
        [(int(name.get("localSheetId")), deepcopy(name)) for name in names if name.get("localSheetId") is not None]
        if names is not None
        else []
    )
    if names is not None:
        for name in list(names):
            scope = name.get("localSheetId")
            if scope is not None:
                index = int(scope)
                if index >= len(session.original_sheets):
                    raise ValueError("Template defined name refers to an invalid worksheet")
                source_sheet = session.original_sheets[index]
                if source_sheet not in current:
                    names.remove(name)
                    changed = True
                    continue
                mapped = str(current.index(source_sheet))
                if mapped != scope:
                    name.set("localSheetId", mapped)
                    changed = True
            keys.add((name.get("name", ""), name.get("localSheetId")))
    for target_sheet in current:
        origin = session.origins.get(target_sheet)
        if origin is None or not origin.clone:
            continue
        source_sheet = _source_sheet(session, origin)
        source_index = session.original_sheets.index(source_sheet)
        for original_index, original_name in original_scoped:
            if original_index != source_index:
                continue
            cloned = deepcopy(original_name)
            cloned.set("localSheetId", str(current.index(target_sheet)))
            cloned.text = relocate_formula_references(
                cloned.text or "", session.original_sheet_names[source_sheet], target_sheet.title
            )
            key = (cloned.get("name", ""), cloned.get("localSheetId"))
            if key not in keys:
                names.append(cloned)
                keys.add(key)
                changed = True
    donor_names = donor_root.find(f"{{{_S}}}definedNames")
    if donor_names is not None:
        for name in donor_names:
            key = (name.get("name", ""), name.get("localSheetId"))
            if key in keys:
                continue
            if names is None:
                names = etree.Element(f"{{{_S}}}definedNames")
                calc = root.find(f"{{{_S}}}calcPr")
                if calc is None:
                    root.append(names)
                else:
                    root.insert(root.index(calc), names)
            names.append(deepcopy(name))
            keys.add(key)
            changed = True
    return changed


def _invalidate_calculation(editor: Any, worksheet_parts: set[str], workbook_root: Any) -> None:
    for part in worksheet_parts:
        root = _parse(editor.parts[part])
        changed = False
        for cell in root.findall(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c"):
            if cell.find(f"{{{_S}}}f") is None:
                continue
            for cached in list(cell):
                if cached.tag in (f"{{{_S}}}v", f"{{{_S}}}is"):
                    cell.remove(cached)
                    changed = True
        if changed:
            editor.parts[part] = _xml(root)
    # Reference caches are derived results; literal chart series are the data.
    # Restrict this to the standard chart part and cache elements we understand.
    editor.finalize_content_types()
    types = _parse(editor.parts["[Content_Types].xml"])
    for declared in types.findall(f"{{{_C}}}Override"):
        if declared.get("ContentType") != "application/vnd.openxmlformats-officedocument.drawingml.chart+xml":
            continue
        part = unquote(declared.get("PartName", "")).lstrip("/")
        if part not in editor.parts:
            continue
        chart = _parse(editor.parts[part])
        changed = False
        for reference_name, cache_name in (
            ("numRef", "numCache"),
            ("strRef", "strCache"),
            ("multiLvlStrRef", "multiLvlStrCache"),
        ):
            for reference in chart.iter(f"{{{_CHART}}}{reference_name}"):
                for cached in list(reference):
                    if cached.tag == f"{{{_CHART}}}{cache_name}":
                        reference.remove(cached)
                        changed = True
        if changed:
            editor.parts[part] = _xml(chart)
    calc = workbook_root.find(f"{{{_S}}}calcPr")
    if calc is None:
        calc = etree.Element(f"{{{_S}}}calcPr")
        later = {
            "oleSize",
            "customWorkbookViews",
            "pivotCaches",
            "smartTagPr",
            "smartTagTypes",
            "webPublishing",
            "fileRecoveryPr",
            "webPublishObjects",
            "extLst",
        }
        following = next((child for child in workbook_root if etree.QName(child).localname in later), None)
        if following is None:
            workbook_root.append(calc)
        else:
            workbook_root.insert(workbook_root.index(following), calc)
    calc.set("fullCalcOnLoad", "1")
    calc.set("forceFullCalc", "1")
    relations_part = "xl/_rels/workbook.xml.rels"
    relationships = _parse(editor.parts[relations_part])
    removed: set[str] = set()
    for relationship in list(relationships):
        if relationship.get("Type", "").endswith("/calcChain"):
            removed.add(_target("xl/workbook.xml", relationship.get("Target", "")))
            relationships.remove(relationship)
    if removed:
        editor.parts[relations_part] = _xml(relationships)
        for part in removed:
            editor.remove_part(part)


def _template_parts(
    workbook: Any, donor: dict[str, bytes], session: TemplateSession, final_states: dict[Any, SheetState]
) -> dict[str, bytes]:
    from .objects import merge_sheet_objects
    from .package import PackageEditor
    from .worksheet import patch_worksheet

    editor = PackageEditor(dict(session.parts))
    styles = StyleMerger(session.styles_xml, donor["xl/styles.xml"], session.styles)
    donor_sheets = _worksheets(donor, all_sheets=True)
    original_root = _parse(session.parts["xl/workbook.xml"])
    donor_root = _parse(donor["xl/workbook.xml"])
    source_nodes = {sheet.get("name"): sheet for sheet in original_root.findall(f"{{{_S}}}sheets/{{{_S}}}sheet")}
    old_sheet_nodes = original_root.find(f"{{{_S}}}sheets")
    next_id = max((int(node.get("sheetId", "0")) for node in source_nodes.values()), default=0) + 1
    output_nodes = []
    output_parts: set[str] = set()
    changed_book = _ordered_sheets(workbook) != session.original_sheets
    calculation_changed = changed_book

    for sheet in _ordered_sheets(workbook):
        regular_sheet = sheet in workbook.worksheets
        if not regular_sheet and sheet in session.original_sheets:
            original_name = session.original_sheet_names[sheet]
            node = deepcopy(source_nodes[original_name])
            if sheet.title != original_name:
                node.set("name", sheet.title)
                changed_book = True
                calculation_changed = True
            if sheet.sheet_state != node.get("state", "visible"):
                node.set("state", sheet.sheet_state)
                changed_book = True
            output_nodes.append(node)
            continue
        origin = session.origins.get(sheet)
        donor_part = donor_sheets[sheet.title]
        donor_sheet_root = _parse(donor[donor_part])
        if origin is None:
            part = editor.import_part(donor, donor_part)
            root = _parse(editor.parts[part])
            _map_sheet_styles(root, styles)
            editor.parts[part] = _xml(root)
            node = etree.Element(f"{{{_S}}}sheet", name=sheet.title, sheetId=str(next_id))
            next_id += 1
            relation_type = "worksheet" if regular_sheet else "chartsheet"
            node.set(f"{{{_R}}}id", editor.add_relationship("xl/workbook.xml", f"{_R}/{relation_type}", part))
            if sheet.sheet_state != "visible":
                node.set("state", sheet.sheet_state)
            changed_book = True
            calculation_changed = True
        else:
            if origin.clone:
                source_sheet = _source_sheet(session, origin)
                part = editor.clone_part(
                    origin.part,
                    source_sheet_name=session.original_sheet_names[source_sheet],
                    target_sheet_name=sheet.title,
                )
                root = _parse(editor.parts[part])
                node = etree.Element(f"{{{_S}}}sheet", name=sheet.title, sheetId=str(next_id))
                next_id += 1
                node.set(f"{{{_R}}}id", editor.add_relationship("xl/workbook.xml", f"{_R}/worksheet", part))
                if sheet.sheet_state != "visible":
                    node.set("state", sheet.sheet_state)
                changed_book = True
            else:
                part = origin.part
                root = _parse(editor.parts[part])
                node = deepcopy(source_nodes[origin.baseline.title])
                if sheet.title != origin.baseline.title:
                    node.set("name", sheet.title)
                    changed_book = True
                    calculation_changed = True
                if sheet.sheet_state != origin.baseline.sheet_state:
                    node.set("state", sheet.sheet_state)
                    changed_book = True
            result = patch_worksheet(
                root, donor_sheet_root, sheet, origin.baseline, styles, final_state=final_states[sheet]
            )
            objects_changed = merge_sheet_objects(
                editor,
                part,
                root,
                donor,
                donor_part,
                donor_sheet_root,
                sheet,
                baseline_images_count=origin.baseline.image_count,
                comment_changes=result.comment_changes,
                hyperlink_changes=result.hyperlink_changes,
            )
            if result.changed or objects_changed:
                editor.parts[part] = _xml(root)
            calculation_changed = calculation_changed or result.calculation_changed
        output_nodes.append(node)
        if regular_sheet:
            output_parts.add(part)

    if changed_book:
        for child in list(old_sheet_nodes):
            old_sheet_nodes.remove(child)
        old_sheet_nodes.extend(output_nodes)
        # Sheet insertion changes the meaning of view indices. Use the final
        # model's view indices while preserving unrelated source view settings.
        views = original_root.find(f"{{{_S}}}bookViews")
        donor_views = donor_root.find(f"{{{_S}}}bookViews")
        if views is not None and donor_views is not None:
            for before, after in zip(views, donor_views, strict=False):
                for key in ("activeTab", "firstSheet"):
                    if key in after.attrib:
                        before.set(key, after.get(key))
    changed_book = _defined_names(original_root, donor_root, session, workbook) or changed_book
    if calculation_changed:
        _invalidate_calculation(editor, output_parts, original_root)
        changed_book = True
    if changed_book:
        editor.parts["xl/workbook.xml"] = _xml(original_root)
    changed_styles = styles.to_bytes()
    if changed_styles is not None:
        editor.parts["xl/styles.xml"] = changed_styles
        if "xl/styles.xml" not in session.parts:
            editor.add_relationship("xl/workbook.xml", f"{_R}/styles", "xl/styles.xml")
            editor.ensure_content_type(
                "xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            )
    editor.deduplicate_new_png_media(set(session.parts))
    editor.finalize_content_types()
    return editor.parts


def save_xlsx_workbook(workbook: Any, output_path: str | Path) -> None:
    """Save explicit edits, retaining untouched template parts and attributes."""
    from .package import PackageEditor

    output = Path(output_path)
    session: TemplateSession | None = getattr(workbook, _SESSION, None)
    if session is not None and output.resolve() == Path(session.source_path):
        raise ValueError("Template output must not overwrite its source")
    check_workbook_text_capacity(workbook)
    # The dependency's writer mutates some in-memory dimensions while saving.
    # Freeze authored state before serialization so those changes are not edits.
    final_states = {sheet: SheetState.capture(sheet) for sheet in workbook.worksheets} if session is not None else {}
    buffer = io.BytesIO()
    workbook.save(buffer)
    donor = _parts(buffer.getvalue())
    _empty_strings(donor, workbook)
    if session is None:
        editor = PackageEditor(donor)
        editor.deduplicate_new_png_media(set())
        editor.finalize_content_types()
        parts = editor.parts
    else:
        parts = _template_parts(workbook, donor, session, final_states)
    descriptor, name = tempfile.mkstemp(prefix="xlsx-edit-", suffix=".xlsx", dir=output.parent)
    os.close(descriptor)
    temporary = Path(name)
    try:
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as package:
            for part, payload in parts.items():
                package.writestr(part, payload)
        os.replace(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)


__all__ = ["load_template_workbook", "register_template_clone", "save_xlsx_workbook"]
