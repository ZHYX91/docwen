"""Apply modeled worksheet edits to the original worksheet XML."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from itertools import pairwise
from typing import Any

from lxml import etree
from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple, get_column_letter, range_boundaries

from .state import CellState, SheetState

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_ORDER = (
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
_PROPERTY_TAGS = {
    "sheet_format": "sheetFormatPr",
    "sheet_properties": "sheetPr",
    "views": "sheetViews",
    "page_margins": "pageMargins",
    "page_setup": "pageSetup",
    "print_options": "printOptions",
}
_EMPTY_CELL = CellState(None, "n", None, None, None)


@dataclass
class PatchResult:
    changed: bool = False
    calculation_changed: bool = False
    comment_changes: set[str] = field(default_factory=set)
    hyperlink_changes: set[str] = field(default_factory=set)


def _tag(name: str) -> str:
    return f"{{{_NS}}}{name}"


def _local(node: Any) -> str:
    return etree.QName(node).localname if isinstance(node.tag, str) else ""


def _insert(root: Any, node: Any) -> None:
    name = _local(node)
    priority = _ORDER.index(name) if name in _ORDER else len(_ORDER)
    for index, existing in enumerate(root):
        other = _local(existing)
        if other in _ORDER and _ORDER.index(other) > priority:
            root.insert(index, node)
            return
    root.append(node)


def _get_or_create(root: Any, name: str) -> Any:
    node = root.find(_tag(name))
    if node is None:
        node = etree.Element(_tag(name))
        _insert(root, node)
    return node


def _row(sheet_data: Any, number: int) -> Any:
    for index, existing in enumerate(sheet_data):
        current = int(existing.get("r", "0"))
        if current == number:
            return existing
        if current > number:
            node = etree.Element(_tag("row"), r=str(number))
            sheet_data.insert(index, node)
            return node
    return etree.SubElement(sheet_data, _tag("row"), r=str(number))


def _cell(sheet_data: Any, address: str) -> Any:
    row_number, column = coordinate_to_tuple(address)
    row = _row(sheet_data, row_number)
    for index, existing in enumerate(row):
        if existing.tag != _tag("c"):
            continue
        if existing.get("r") == address:
            return existing
        if coordinate_to_tuple(existing.get("r"))[1] > column:
            node = etree.Element(_tag("c"), r=address)
            row.insert(index, node)
            return node
    return etree.SubElement(row, _tag("c"), r=address)


def _set(node: Any, name: str, value: Any) -> None:
    if value is None:
        node.attrib.pop(name, None)
    elif isinstance(value, bool):
        node.set(name, "1" if value else "0")
    else:
        node.set(name, str(value))


def _same_style(left: Any, right: Any) -> bool:
    return (left or (0,) * 9) == (right or (0,) * 9)


def _write_value(cell: Any, donor: Any, state: CellState) -> None:
    for child in tuple(cell):
        if child.tag in {_tag("f"), _tag("v"), _tag("is")}:
            cell.remove(child)
    cell.attrib.pop("t", None)
    values = []
    if state.value == "" and state.data_type == "s":
        cell.set("t", "inlineStr")
        inline = etree.Element(_tag("is"))
        text = etree.SubElement(inline, _tag("t"))
        text.set(_XML_SPACE, "preserve")
        text.text = ""
        values.append(inline)
    elif donor is not None:
        if donor.get("t") is not None:
            cell.set("t", donor.get("t"))
        values.extend(deepcopy(child) for child in donor if child.tag in {_tag("f"), _tag("v"), _tag("is")})
    elif state.value is not None:
        raise ValueError(f"Changed XLSX cell {cell.get('r')} is absent from the generated worksheet")
    for index, child in enumerate(values):
        cell.insert(index, child)


def _patch_cells(
    root: Any, donor_root: Any, before: SheetState, after: SheetState, styles: Any, result: PatchResult
) -> None:
    donor_cells = {cell.get("r"): cell for cell in donor_root.findall(f"{_tag('sheetData')}/{_tag('row')}/{_tag('c')}")}
    addresses = set(before.cells) | set(after.cells)
    sheet_data = root.find(_tag("sheetData"))
    for address in sorted(addresses, key=coordinate_to_tuple):
        old = before.cells.get(address, _EMPTY_CELL)
        new = after.cells.get(address, _EMPTY_CELL)
        if old.comment != new.comment:
            result.comment_changes.add(address)
        if old.hyperlink != new.hyperlink:
            result.hyperlink_changes.add(address)
        value_changed = (old.value, old.data_type) != (new.value, new.data_type)
        style_changed = not _same_style(old.style, new.style)
        if not value_changed and not style_changed:
            continue
        if sheet_data is None:
            sheet_data = _get_or_create(root, "sheetData")
        cell = _cell(sheet_data, address)
        donor = donor_cells.get(address)
        if value_changed:
            _write_value(cell, donor, new)
            result.calculation_changed = True
        if style_changed:
            donor_id = int(donor.get("s", "0")) if donor is not None else 0
            mapped = styles.map_style(donor_id, source_style_id=int(cell.get("s", "0")))
            _set(cell, "s", mapped if mapped else None)
        result.changed = True


def _expand_changed_shared_formulas(
    root: Any, donor_root: Any, before: SheetState, after: SheetState, result: PatchResult
) -> None:
    """Detach an edited shared group before any of its formula owners disappear."""
    changed = {
        address
        for address in before.cells.keys() | after.cells.keys()
        if (
            before.cells.get(address, _EMPTY_CELL).value,
            before.cells.get(address, _EMPTY_CELL).data_type,
        )
        != (
            after.cells.get(address, _EMPTY_CELL).value,
            after.cells.get(address, _EMPTY_CELL).data_type,
        )
    }
    if not changed:
        return
    groups: dict[str | None, list[tuple[str, Any, Any]]] = {}
    for cell in root.findall(f"{_tag('sheetData')}/{_tag('row')}/{_tag('c')}"):
        formula = cell.find(_tag("f"))
        if formula is not None and formula.get("t") == "shared":
            address = cell.get("r", "")
            groups.setdefault(formula.get("si"), []).append((address, cell, formula))
    donors = {cell.get("r"): cell for cell in donor_root.findall(f"{_tag('sheetData')}/{_tag('row')}/{_tag('c')}")}
    expanded = []
    for identity, members in groups.items():
        if not any(address in changed for address, _cell_node, _formula in members):
            continue
        masters = [formula for _address, _cell_node, formula in members if formula.text]
        if identity is None or len(masters) != 1 or not masters[0].get("ref"):
            raise ValueError("An edited XLSX shared formula group has no unique master and range")
        left, top, right, bottom = _bounds(masters[0].get("ref"))
        addresses = {address for address, _cell_node, _formula in members}
        if len(addresses) != len(members):
            raise ValueError("An edited XLSX shared formula group contains duplicate cell owners")
        for address, cell, formula in members:
            row, column = coordinate_to_tuple(address)
            if not left <= column <= right or not top <= row <= bottom:
                raise ValueError("An edited XLSX shared formula member lies outside its master range")
            if address in changed:
                # The regular cell writer supplies the newly authored value or formula.
                continue
            state = after.cells.get(address, _EMPTY_CELL)
            donor = donors.get(address)
            donor_formula = donor.find(_tag("f")) if donor is not None else None
            if (
                state.data_type != "f"
                or not isinstance(state.value, str)
                or not state.value.startswith("=")
                or donor_formula is None
                or donor_formula.get("t") not in (None, "normal")
                or not donor_formula.text
                or donor_formula.text != state.value[1:]
            ):
                raise ValueError("An edited XLSX shared formula cannot be safely expanded")
            expanded.append((cell, formula, donor_formula.text))
    for cell, formula, text in expanded:
        for attribute in ("t", "si", "ref"):
            formula.attrib.pop(attribute, None)
        formula.text = text
        for cached in tuple(cell.findall(_tag("v"))):
            cell.remove(cached)
        result.changed = True
        result.calculation_changed = True


def _different(old: Any, new: Any, name: str) -> bool:
    if name == "style":
        return not _same_style(old, new)
    if name in {"hidden", "outlineLevel", "collapsed", "thickTop", "thickBot", "bestFit"}:
        return (old or 0) != (new or 0)
    return old != new


def _patch_rows(
    root: Any, donor_root: Any, before: SheetState, after: SheetState, styles: Any, result: PatchResult
) -> None:
    donors = {int(row.get("r")): row for row in donor_root.findall(f"{_tag('sheetData')}/{_tag('row')}")}
    for number in sorted(set(before.rows) | set(after.rows)):
        old, new = before.rows.get(number, {}), after.rows.get(number, {})
        fields = [name for name in old.keys() | new.keys() if _different(old.get(name), new.get(name), name)]
        if not fields:
            continue
        row = _row(_get_or_create(root, "sheetData"), number)
        donor = donors.get(number)
        for name in fields:
            value = new.get(name)
            if name == "height":
                _set(row, "ht", value)
                _set(row, "customHeight", True if value is not None else None)
            elif name == "style":
                donor_id = int(donor.get("s", "0")) if donor is not None else 0
                mapped = styles.map_style(donor_id, source_style_id=int(row.get("s", "0")))
                _set(row, "s", mapped if mapped else None)
                _set(row, "customFormat", True if mapped else None)
            else:
                _set(row, name, value)
        result.changed = True


def _column_range(key: str, data: dict[str, Any]) -> tuple[int, int]:
    start = data.get("min") or column_index_from_string(key)
    return start, data.get("max") or start


def _bounds(reference: str) -> tuple[int, int, int, int]:
    left, top, right, bottom = range_boundaries(reference)
    if left is None or top is None or right is None or bottom is None:
        raise ValueError("XLSX worksheet dimensions must have finite cell bounds")
    return left, top, right, bottom


def _column_segments(cols: Any, start: int, end: int) -> list[Any]:
    """Split original column ranges only where a modeled edit takes effect."""
    overlapping = [item for item in cols if int(item.get("min", "0")) <= end and int(item.get("max", "0")) >= start]
    boundaries = {start, end + 1}
    for item in overlapping:
        boundaries.add(max(start, int(item.get("min"))))
        boundaries.add(min(end + 1, int(item.get("max")) + 1))
    segments = []
    points = sorted(boundaries)
    for left, right in pairwise(points):
        source = next(
            (item for item in reversed(overlapping) if int(item.get("min")) <= left <= int(item.get("max"))), None
        )
        segment = deepcopy(source) if source is not None else etree.Element(_tag("col"))
        segment.set("min", str(left))
        segment.set("max", str(right - 1))
        segments.append(segment)
    outside = []
    for item in overlapping:
        if int(item.get("min")) < start:
            left = deepcopy(item)
            left.set("max", str(start - 1))
            outside.append(left)
        if int(item.get("max")) > end:
            right = deepcopy(item)
            right.set("min", str(end + 1))
            outside.append(right)
        cols.remove(item)
    for item in outside + segments:
        position = next(
            (index for index, existing in enumerate(cols) if int(existing.get("min", "0")) > int(item.get("min"))),
            len(cols),
        )
        cols.insert(position, item)
    return segments


def _patch_columns(
    root: Any, donor_root: Any, before: SheetState, after: SheetState, styles: Any, result: PatchResult
) -> None:
    donor_columns = donor_root.findall(f"{_tag('cols')}/{_tag('col')}")
    for key in sorted(set(before.columns) | set(after.columns), key=column_index_from_string):
        old, new = before.columns.get(key, {}), after.columns.get(key, {})
        fields = [name for name in old.keys() | new.keys() if _different(old.get(name), new.get(name), name)]
        if not fields:
            continue
        start, end = _column_range(key, new or old)
        cols = _get_or_create(root, "cols")
        donor = next(
            (item for item in donor_columns if int(item.get("min", "0")) == start and int(item.get("max", "0")) == end),
            None,
        )
        for column in _column_segments(cols, start, end):
            for name in fields:
                value = new.get(name)
                if name == "style":
                    donor_id = int(donor.get("style", "0")) if donor is not None else 0
                    mapped = styles.map_style(donor_id, source_style_id=int(column.get("style", "0")))
                    _set(column, "style", mapped if mapped else None)
                elif name == "width":
                    _set(column, "width", value)
                    _set(column, "customWidth", True if value is not None else None)
                elif name not in {"min", "max"}:
                    _set(column, name, value)
        result.changed = True


def _model_patch(original: Any, before: Any, after: Any) -> None:
    """Copy changed modeled properties while retaining unmodeled attributes/children."""
    old_attrs = before.attrib if before is not None else {}
    new_attrs = after.attrib if after is not None else {}
    for name in set(old_attrs) | set(new_attrs):
        if old_attrs.get(name) != new_attrs.get(name):
            _set(original, name, new_attrs.get(name))
    if before is not None and after is not None and before.text != after.text:
        original.text = after.text
    old_children: dict[str, list[Any]] = {}
    new_children: dict[str, list[Any]] = {}
    for source, target in ((before, old_children), (after, new_children)):
        if source is not None:
            for child in source:
                target.setdefault(_local(child), []).append(child)
    for name in old_children.keys() | new_children.keys():
        old_items, new_items = old_children.get(name, []), new_children.get(name, [])
        originals = [item for item in original if _local(item) == name and etree.QName(item).namespace == _NS]
        for index in range(max(len(old_items), len(new_items))):
            old = old_items[index] if index < len(old_items) else None
            new = new_items[index] if index < len(new_items) else None
            target = originals[index] if index < len(originals) else None
            if new is None:
                if target is not None:
                    original.remove(target)
            elif target is not None:
                _model_patch(target, old, new)
            else:
                node = deepcopy(new)
                for descendant in node.iter():
                    if isinstance(descendant.tag, str) and not etree.QName(descendant).namespace:
                        descendant.tag = _tag(_local(descendant))
                original.append(node)


def _patch_properties(root: Any, before: SheetState, after: SheetState, result: PatchResult) -> None:
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    for name, tag in _PROPERTY_TAGS.items():
        old, new = before.properties.get(name), after.properties.get(name)
        if old == new:
            continue
        original = root.find(_tag(tag))
        if original is None:
            original = _get_or_create(root, tag)
        _model_patch(
            original,
            etree.fromstring(old, parser) if old is not None else None,
            etree.fromstring(new, parser) if new is not None else None,
        )
        result.changed = True


def _patch_merges(root: Any, before: SheetState, after: SheetState, result: PatchResult) -> None:
    if before.merges == after.merges:
        return
    merges = _get_or_create(root, "mergeCells")
    for item in tuple(merges):
        if item.tag == _tag("mergeCell") and item.get("ref") in before.merges - after.merges:
            merges.remove(item)
    for address in sorted(after.merges - before.merges):
        etree.SubElement(merges, _tag("mergeCell"), ref=address)
    count = len(merges.findall(_tag("mergeCell")))
    if count:
        merges.set("count", str(count))
    elif not len(merges) and not (set(merges.attrib) - {"count"}):
        root.remove(merges)
    else:
        merges.set("count", "0")
    result.changed = True
    result.calculation_changed = True


def patch_worksheet(
    root: Any,
    donor_root: Any,
    worksheet: Any,
    baseline: SheetState,
    styles: Any,
    *,
    final_state: SheetState | None = None,
) -> PatchResult:
    """Apply only operations visible between the post-load baseline and final model."""
    result = PatchResult()
    final = final_state if final_state is not None else SheetState.capture(worksheet)
    _expand_changed_shared_formulas(root, donor_root, baseline, final, result)
    _patch_cells(root, donor_root, baseline, final, styles, result)
    _patch_rows(root, donor_root, baseline, final, styles, result)
    _patch_columns(root, donor_root, baseline, final, styles, result)
    _patch_properties(root, baseline, final, result)
    _patch_merges(root, baseline, final, result)
    if set(final.cells) - set(baseline.cells):
        donor_dimension = donor_root.find(_tag("dimension"))
        if donor_dimension is not None:
            dimension = _get_or_create(root, "dimension")
            donor_reference = donor_dimension.get("ref")
            original_reference = dimension.get("ref")
            reference = donor_reference
            if original_reference is not None:
                original_bounds = _bounds(original_reference)
                donor_bounds = _bounds(donor_reference)
                left = min(original_bounds[0], donor_bounds[0])
                top = min(original_bounds[1], donor_bounds[1])
                right = max(original_bounds[2], donor_bounds[2])
                bottom = max(original_bounds[3], donor_bounds[3])
                reference = f"{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}"
                if original_bounds == (left, top, right, bottom):
                    reference = original_reference
            if original_reference != reference:
                dimension.set("ref", reference)
                result.changed = True
    return result
