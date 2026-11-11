"""Append edited styles while keeping the template's authored style records."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import Any

from lxml import etree
from openpyxl.styles.stylesheet import write_stylesheet

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_ORDER = (
    "numFmts",
    "fonts",
    "fills",
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
)
_COMPONENTS = {"fontId": "fonts", "fillId": "fills", "borderId": "borders"}


def _parse(payload: bytes) -> Any:
    return etree.fromstring(payload, parser=etree.XMLParser(resolve_entities=False, no_network=True))


def _name(element: Any) -> str:
    name = str(element.tag)
    return name.removeprefix("{" + _NS + "}")


def _fingerprint(element: Any) -> tuple[Any, ...]:
    """Compare descriptor output without depending on its namespace prefixes."""
    return (
        _name(element),
        tuple(sorted(element.attrib.items())),
        element.text or "",
        tuple(_fingerprint(child) for child in element),
    )


def _groups(root: Any) -> dict[str, list[Any]]:
    return {_name(child): list(child) for child in root}


@dataclass(frozen=True)
class NormalizedStyleState:
    """A style-only snapshot; capturing it never saves or consumes drawings."""

    xml: bytes


def capture_style_state(workbook: Any) -> NormalizedStyleState:
    return NormalizedStyleState(etree.tostring(write_stylesheet(workbook)))


class StyleMerger:
    """Map donor style IDs into an append-only copy of the original stylesheet.

    The baseline is the same workbook's state immediately after loading. It
    establishes which normalized donor entries refer to original records,
    including records carrying attributes openpyxl does not understand.
    """

    def __init__(self, source_xml: bytes, donor_xml: bytes, baseline: NormalizedStyleState) -> None:
        self._root = _parse(source_xml)
        self._source = _groups(self._root)
        self._original_counts = {name: len(items) for name, items in self._source.items()}
        self._donor = _groups(_parse(donor_xml))
        self._baseline = _groups(_parse(baseline.xml))
        self._changed = False
        self._mapped: dict[tuple[str, int], int] = {}
        self._styles: dict[tuple[int, int | None], int] = {}
        self._formats: dict[int, int] = {}
        self._format_by_code = {
            item.get("formatCode", ""): int(item.get("numFmtId")) for item in self._source.get("numFmts", [])
        }
        self._donor_formats = {int(item.get("numFmtId")): item for item in self._donor.get("numFmts", [])}
        self._reserved_formats = set(self._format_by_code.values())

    def _container(self, name: str) -> Any:
        found = self._root.find(f"{{{_NS}}}{name}")
        if found is not None:
            return found
        node = etree.Element(f"{{{_NS}}}{name}")
        rank = _ORDER.index(name)
        for index, child in enumerate(self._root):
            child_name = _name(child)
            if child_name in _ORDER and _ORDER.index(child_name) > rank:
                self._root.insert(index, node)
                break
        else:
            self._root.append(node)
        self._source[name] = []
        return node

    def _append(self, group: str, element: Any) -> int:
        container = self._container(group)
        members = self._source[group]
        index = len(members)
        container.append(element)
        members.append(element)
        container.set("count", str(len(members)))
        self._changed = True
        return index

    def _original_index(self, group: str, baseline_index: int) -> int | None:
        if group == "cellStyleXfs":
            # Named-style loading may remove duplicate/unused references.
            names = {
                node.get("name")
                for node in self._baseline.get("cellStyles", [])
                if int(node.get("xfId", "0")) == baseline_index
            }
            ids = {
                int(node.get("xfId", "0")) for node in self._source.get("cellStyles", []) if node.get("name") in names
            }
            if len(ids) == 1:
                return ids.pop()
            return None
        if baseline_index < self._original_counts.get(group, 0):
            return baseline_index
        return None

    def _existing(self, group: str, element: Any, preferred_index: int | None = None) -> int | None:
        fingerprint = _fingerprint(element)
        candidates = self._baseline.get(group, [])
        if (
            preferred_index is not None
            and 0 <= preferred_index < len(candidates)
            and _fingerprint(candidates[preferred_index]) == fingerprint
        ):
            original_index = self._original_index(group, preferred_index)
            if original_index is not None:
                return original_index
        for index, candidate in enumerate(candidates):
            if _fingerprint(candidate) == fingerprint:
                original_index = self._original_index(group, index)
                if original_index is not None:
                    return original_index
        return None

    def _map_format(self, number_format_id: int) -> int:
        if number_format_id < 164:
            return number_format_id
        if number_format_id in self._formats:
            return self._formats[number_format_id]
        donor = self._donor_formats.get(number_format_id)
        if donor is None:
            raise ValueError(f"Donor style references missing number format {number_format_id}")
        code = donor.get("formatCode", "")
        if code in self._format_by_code:
            result = self._format_by_code[code]
        else:
            result = max({163, *self._reserved_formats}) + 1
            node = deepcopy(donor)
            node.set("numFmtId", str(result))
            self._append("numFmts", node)
            self._format_by_code[code] = result
            self._reserved_formats.add(result)
        self._formats[number_format_id] = result
        return result

    def _map_attribute(self, name: str, value: str) -> str:
        if name in _COMPONENTS:
            return str(self._map_component(_COMPONENTS[name], int(value)))
        if name == "numFmtId":
            return str(self._map_format(int(value)))
        if name == "xfId":
            index = int(value)
            if index >= len(self._donor.get("cellStyleXfs", [])):
                # Existing cell XFs can retain an original unused style-XF ID
                # although openpyxl only serializes referenced named styles.
                if index < self._original_counts.get("cellStyleXfs", 0):
                    return value
                raise ValueError(f"Donor style references missing style XF {index}")
            return str(self._map_component("cellStyleXfs", index))
        return value

    def _map_component(self, group: str, donor_index: int) -> int:
        key = (group, donor_index)
        if key in self._mapped:
            return self._mapped[key]
        donor_items = self._donor.get(group, [])
        if donor_index < 0 or donor_index >= len(donor_items):
            raise ValueError(f"Donor style references missing {group} entry {donor_index}")
        donor = donor_items[donor_index]
        existing = self._existing(group, donor, donor_index)
        if existing is not None:
            self._mapped[key] = existing
            if group == "cellStyleXfs":
                self._register_named_styles(donor_index, existing)
            return existing
        candidate = deepcopy(donor)
        if group == "cellStyleXfs":
            for attribute in tuple(candidate.attrib):
                if attribute != "xfId":
                    candidate.set(attribute, self._map_attribute(attribute, candidate.get(attribute)))
        fingerprint = _fingerprint(candidate)
        for index, node in enumerate(self._source.get(group, [])):
            if _fingerprint(node) == fingerprint:
                self._mapped[key] = index
                if group == "cellStyleXfs":
                    self._register_named_styles(donor_index, index)
                return index
        result = self._append(group, candidate)
        self._mapped[key] = result
        if group == "cellStyleXfs":
            self._register_named_styles(donor_index, result)
        return result

    def _register_named_styles(self, donor_index: int, source_index: int) -> None:
        original_names = {node.get("name") for node in self._baseline.get("cellStyles", [])}
        current_names = {node.get("name") for node in self._source.get("cellStyles", [])}
        for node in self._donor.get("cellStyles", []):
            name = node.get("name")
            if int(node.get("xfId", "0")) != donor_index or name in original_names or name in current_names:
                continue
            added = deepcopy(node)
            added.set("xfId", str(source_index))
            self._append("cellStyles", added)
            current_names.add(name)

    def _overlay(self, original: Any, before: Any, after: Any, *, xf: bool = False) -> Any:
        """Apply only descriptor changes, retaining unknown source metadata."""
        result = deepcopy(original)
        for name in set(before.attrib) | set(after.attrib):
            if before.get(name) == after.get(name):
                continue
            if name in after.attrib:
                value = after.get(name)
                result.set(name, self._map_attribute(name, value) if xf else value)
            else:
                result.attrib.pop(name, None)
        old_children = {_name(node): node for node in before}
        new_children = {_name(node): node for node in after}
        for name in old_children.keys() | new_children.keys():
            old = old_children.get(name)
            new = new_children.get(name)
            if old is not None and new is not None and _fingerprint(old) == _fingerprint(new):
                continue
            source = next((node for node in result if _name(node) == name), None)
            if new is None:
                if source is not None:
                    result.remove(source)
            elif source is not None and old is not None:
                result.replace(source, self._overlay(source, old, new))
            elif source is not None:
                result.replace(source, deepcopy(new))
            else:
                result.append(deepcopy(new))
        return result

    def map_style(self, donor_style_id: int, *, source_style_id: int | None = None) -> int:
        """Map one edited cell/row/column style, optionally retaining its lineage."""
        key = (donor_style_id, source_style_id)
        if key in self._styles:
            return self._styles[key]
        donors = self._donor.get("cellXfs", [])
        if donor_style_id < 0 or donor_style_id >= len(donors):
            raise ValueError(f"Donor references missing cell style {donor_style_id}")
        donor = donors[donor_style_id]
        if source_style_id is not None:
            if not 0 <= source_style_id < self._original_counts.get("cellXfs", 0):
                raise ValueError(f"Template references missing cell style {source_style_id}")
            before = self._baseline.get("cellXfs", [])[source_style_id]
            original = self._source["cellXfs"][source_style_id]
            if _fingerprint(before) == _fingerprint(donor):
                self._styles[key] = source_style_id
                return source_style_id
            candidate = self._overlay(original, before, donor, xf=True)
        else:
            existing = self._existing("cellXfs", donor, donor_style_id)
            if existing is not None:
                self._styles[key] = existing
                return existing
            candidate = deepcopy(donor)
            for attribute in tuple(candidate.attrib):
                candidate.set(attribute, self._map_attribute(attribute, candidate.get(attribute)))
        fingerprint = _fingerprint(candidate)
        for index, node in enumerate(self._source.get("cellXfs", [])):
            if _fingerprint(node) == fingerprint:
                self._styles[key] = index
                return index
        result = self._append("cellXfs", candidate)
        self._styles[key] = result
        return result

    def to_bytes(self) -> bytes | None:
        if not self._changed:
            return None
        return etree.tostring(self._root, encoding="utf-8", xml_declaration=True)
