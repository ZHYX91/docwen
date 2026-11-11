"""OPC package edits that retain unrelated template parts and relationships."""

from __future__ import annotations

import hashlib
import posixpath
import re
from collections.abc import Callable, Mapping
from typing import Any
from urllib.parse import quote, unquote, urlsplit

from lxml import etree

PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
RELATIONSHIP_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml"
TABLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
_FORMULA_CONTENT_TYPES = {
    TABLE_CONTENT_TYPE,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
}
_FORMULA_TAGS = {
    f"{{{SHEET_NS}}}f",
    f"{{{SHEET_NS}}}calculatedColumnFormula",
    f"{{{SHEET_NS}}}totalsRowFormula",
    f"{{{CHART_NS}}}f",
}
_ERROR_LITERAL = re.compile(r"#(?:N/A(?![\w.])|DIV/0!|GETTING_DATA(?![\w.])|[A-Z][A-Z0-9_.]*[!?])", re.IGNORECASE)


def _rewrite_range_tokens(formula: str, replace: Callable[[str], str]) -> str:
    """Visit reference operands without parsing or normalizing the expression.

    Excel's implicit-intersection and spill operators are valid even though
    openpyxl's formula tokenizer cannot parse all of them. Only operand bounds
    matter here: quoted strings are opaque, as are delimiters inside sheet
    quotes or structured-reference brackets. Preserve all other text exactly.
    """
    delimiters = "=+-*/^&<>()%,;{}@#"
    fragments: list[str] = []
    index = 0
    while index < len(formula):
        start = index
        character = formula[index]
        if character == '"':
            index = _quoted_end(formula, index, '"')
            fragments.append(formula[start:index])
            continue
        error = _ERROR_LITERAL.match(formula, index) if character == "#" else None
        if error is not None:
            # Error literals such as #N/A must not expose N as a table name.
            index = error.end()
            fragments.append(formula[start:index])
            continue
        if character.isspace() or character in delimiters:
            fragments.append(character)
            index += 1
            continue
        brackets = 0
        while index < len(formula):
            character = formula[index]
            if brackets:
                if character == "'" and index + 1 < len(formula):
                    # Structured headers escape special characters with '.
                    index += 2
                    continue
                if character == "[":
                    brackets += 1
                elif character == "]":
                    brackets -= 1
            elif character == "'":
                index = _quoted_end(formula, index, "'")
                continue
            elif character == "[":
                brackets = 1
            elif character == '"' or character.isspace() or character in delimiters:
                break
            index += 1
        reference = formula[start:index]
        # Keep bare function identifiers; qualified local LAMBDA names still
        # need their sheet qualifier relocated when the definition is cloned.
        function = formula[index : index + 1] == "(" and "!" not in reference
        fragments.append(reference if function else replace(reference))
    return "".join(fragments)


def _quoted_end(formula: str, start: int, quote_character: str) -> int:
    index = start + 1
    while index < len(formula):
        if formula[index] == quote_character:
            if formula[index : index + 2] == quote_character * 2:
                index += 2
                continue
            return index + 1
        index += 1
    return index


def _range_segments(reference: str) -> list[str]:
    segments: list[str] = []
    start = 0
    quoted = False
    brackets = 0
    index = 0
    while index < len(reference):
        character = reference[index]
        if character == "'":
            if quoted and reference[index : index + 2] == "''":
                index += 2
                continue
            quoted = not quoted
        elif not quoted:
            if character == "[":
                brackets += 1
            elif character == "]":
                brackets -= 1
            elif character == ":" and brackets == 0:
                segments.append(reference[start:index])
                start = index + 1
        index += 1
    segments.append(reference[start:])
    return segments


def relocate_formula_references(formula: str, source: str, target: str) -> str:
    """Retarget explicit local sheet references while retaining literals and external/3D references."""
    if source == target:
        return formula
    folded = formula.casefold()
    quoted_source = "'" + source.replace("'", "''") + "'"
    if source.casefold() + "!" not in folded and quoted_source.casefold() + "!" not in folded:
        return formula
    target_qualifier = "'" + target.replace("'", "''") + "'"

    def replace_one(reference: str) -> str:
        qualifier, separator, address = reference.rpartition("!")
        if not separator:
            return reference
        if qualifier.startswith("'") and qualifier.endswith("'"):
            sheet_name = qualifier[1:-1].replace("''", "'")
        elif "'" in qualifier:
            return reference
        else:
            sheet_name = qualifier
        if any(character in sheet_name for character in "[]:") or sheet_name.casefold() != source.casefold():
            return reference
        return target_qualifier + "!" + address

    def replace(reference: str) -> str:
        segments = _range_segments(reference)
        if len(segments) == 2 and all("!" in segment for segment in segments):
            return ":".join(replace_one(segment) for segment in segments)
        return replace_one(reference)

    return _rewrite_range_tokens(formula, replace)


def _clear_formula_cache(formula: Any) -> None:
    parent = formula.getparent()
    if parent is None:
        return
    if parent.tag == f"{{{SHEET_NS}}}c":
        for cached in parent.findall(f"{{{SHEET_NS}}}v"):
            parent.remove(cached)
    elif formula.tag == f"{{{CHART_NS}}}f":
        for cache_name in ("numCache", "strCache", "multiLvlStrCache"):
            for cached in parent.findall(f"{{{CHART_NS}}}{cache_name}"):
                parent.remove(cached)


def parse_xml(payload: bytes) -> Any:
    return etree.fromstring(payload, parser=etree.XMLParser(resolve_entities=False, no_network=True))


def serialize_xml(root: Any) -> bytes:
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True)


def _part_name(part: str) -> str:
    if not part or "\\" in part or part.startswith("/") or any(c in part for c in "\x00?#"):
        raise ValueError(f"Invalid package part name: {part!r}")
    normalized = posixpath.normpath(part)
    if normalized in {".", ".."} or normalized.startswith("../") or normalized != part:
        raise ValueError(f"Package part is outside the package: {part!r}")
    return normalized


def rels_part(owner_part: str) -> str:
    if not owner_part:
        return "_rels/.rels"
    directory, filename = posixpath.split(_part_name(owner_part))
    return posixpath.join(directory, "_rels", filename + ".rels")


def resolve_target(owner_part: str, target: str) -> str:
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc or parsed.query or "\\" in target:
        raise ValueError(f"Invalid internal relationship target: {target!r}")
    path = unquote(parsed.path)
    if not path:
        if parsed.fragment and owner_part:
            return _part_name(owner_part)
        raise ValueError("An internal relationship has an empty target")
    directory = posixpath.dirname(_part_name(owner_part)) if owner_part else ""
    resolved = posixpath.normpath(path.lstrip("/") if path.startswith("/") else posixpath.join(directory, path))
    return _part_name(resolved)


def _relative_target(owner_part: str, target: str) -> str:
    directory = posixpath.dirname(_part_name(owner_part)) if owner_part else ""
    relative = posixpath.relpath(_part_name(target), directory or ".")
    return quote(relative, safe="/!$&'()*+,-.:;=@_~[]")


def _types_root(parts: Mapping[str, bytes]) -> Any:
    if "[Content_Types].xml" not in parts:
        raise ValueError("XLSX package is missing [Content_Types].xml")
    root = parse_xml(parts["[Content_Types].xml"])
    if root.tag != f"{{{CONTENT_TYPES_NS}}}Types":
        raise ValueError("Invalid XLSX content-types document")
    return root


def _content_type(root: Any, part: str) -> str:
    overrides = [
        node.get("ContentType", "")
        for node in root.findall(f"{{{CONTENT_TYPES_NS}}}Override")
        if unquote(node.get("PartName", "")).lstrip("/") == part
    ]
    extension = posixpath.splitext(part)[1].removeprefix(".").lower()
    matches = overrides or [
        node.get("ContentType", "")
        for node in root.findall(f"{{{CONTENT_TYPES_NS}}}Default")
        if node.get("Extension", "").lower() == extension
    ]
    if not matches or not matches[0] or len(set(matches)) != 1:
        raise ValueError(f"Missing or conflicting content type for {part!r}")
    return matches[0]


def _owner_of_rels(part: str) -> str:
    if part == "_rels/.rels":
        return ""
    directory, filename = posixpath.split(part)
    if posixpath.basename(directory) != "_rels" or not filename.endswith(".rels"):
        raise ValueError(f"Invalid relationships part name: {part!r}")
    return posixpath.join(posixpath.dirname(directory), filename[:-5])


class PackageEditor:
    """Preserve the original package, importing only explicitly requested graphs."""

    def __init__(self, source: Mapping[str, bytes]) -> None:
        self.parts = dict(source)
        self._content_types = _types_root(self.parts)
        self._content_types_changed = False
        self._reserved = set(self.parts)

    def allocate(self, preferred_part: str) -> str:
        preferred = _part_name(preferred_part)
        stem, extension = posixpath.splitext(preferred)
        stem = re.sub(r"_copy\d+$", "", stem)
        candidate = preferred
        index = 0
        while candidate in self.parts or candidate in self._reserved or rels_part(candidate) in self.parts:
            index += 1
            candidate = f"{stem}_copy{index}{extension}"
        self._reserved.add(candidate)
        return candidate

    def relationships(self, owner_part: str) -> Any:
        payload = self.parts.get(rels_part(owner_part))
        if payload is None:
            return parse_xml(f'<Relationships xmlns="{PACKAGE_REL_NS}"/>'.encode())
        root = parse_xml(payload)
        if root.tag != f"{{{PACKAGE_REL_NS}}}Relationships":
            raise ValueError(f"Invalid relationships document for {owner_part!r}")
        return root

    def save_relationships(self, owner_part: str, root: Any) -> None:
        if root.tag != f"{{{PACKAGE_REL_NS}}}Relationships":
            raise ValueError("Expected an OPC Relationships element")
        part = rels_part(owner_part)
        self.parts[part] = serialize_xml(root)
        self._reserved.add(part)
        self.ensure_content_type(part, RELATIONSHIP_CONTENT_TYPE)

    def add_relationship(
        self, owner_part: str, relationship_type: str, target: str, target_mode: str | None = None
    ) -> str:
        if target_mode not in {None, "Internal", "External"}:
            raise ValueError(f"Unsupported relationship target mode: {target_mode!r}")
        root = self.relationships(owner_part)
        identifiers = {node.get("Id") for node in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship")}
        index = 1
        while f"rId{index}" in identifiers:
            index += 1
        identifier = f"rId{index}"
        attributes = {
            "Id": identifier,
            "Type": relationship_type,
            "Target": target if target_mode == "External" else _relative_target(owner_part, target),
        }
        if target_mode is not None:
            attributes["TargetMode"] = target_mode
        etree.SubElement(root, f"{{{PACKAGE_REL_NS}}}Relationship", attributes)
        self.save_relationships(owner_part, root)
        return identifier

    def relationship_target(self, owner_part: str, identifier: str) -> tuple[str, str, str | None]:
        matches = [
            node
            for node in self.relationships(owner_part).findall(f"{{{PACKAGE_REL_NS}}}Relationship")
            if node.get("Id") == identifier
        ]
        if len(matches) != 1:
            raise ValueError(f"Missing or ambiguous relationship {identifier!r} in {owner_part!r}")
        relationship = matches[0]
        target = relationship.get("Target", "")
        mode = relationship.get("TargetMode")
        if mode not in {None, "Internal", "External"}:
            raise ValueError(f"Unsupported relationship target mode: {mode!r}")
        if mode != "External":
            target = resolve_target(owner_part, target)
            if target not in self.parts:
                raise ValueError(f"Relationship target is missing: {target!r}")
            mode = None
        return target, relationship.get("Type", ""), mode

    def ensure_content_type(self, part: str, content_type: str) -> None:
        part = _part_name(part)
        if not content_type:
            raise ValueError(f"Empty content type for {part!r}")
        overrides = [
            node
            for node in self._content_types.findall(f"{{{CONTENT_TYPES_NS}}}Override")
            if unquote(node.get("PartName", "")).lstrip("/") == part
        ]
        if overrides:
            if any(node.get("ContentType") != content_type for node in overrides):
                raise ValueError(f"Content type collision for {part!r}")
            return
        extension = posixpath.splitext(part)[1].removeprefix(".").lower()
        defaults = [
            node
            for node in self._content_types.findall(f"{{{CONTENT_TYPES_NS}}}Default")
            if node.get("Extension", "").lower() == extension
        ]
        if defaults and all(node.get("ContentType") == content_type for node in defaults):
            return
        etree.SubElement(
            self._content_types,
            f"{{{CONTENT_TYPES_NS}}}Override",
            PartName="/" + quote(part, safe="/!$&'()*+,-.:;=@_~[]"),
            ContentType=content_type,
        )
        self._content_types_changed = True

    def _table_registry(self) -> tuple[int, set[str]]:
        maximum = 0
        names: set[str] = set()
        for node in self._content_types.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
            part = unquote(node.get("PartName", "")).lstrip("/")
            if part not in self.parts:
                continue
            content_type = node.get("ContentType", "")
            if content_type in {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                "application/vnd.ms-excel.template.macroEnabled.main+xml",
            }:
                workbook = parse_xml(self.parts[part])
                names.update(
                    name.casefold()
                    for defined in workbook.findall(f"{{{SHEET_NS}}}definedNames/{{{SHEET_NS}}}definedName")
                    if (name := defined.get("name"))
                )
            if content_type != TABLE_CONTENT_TYPE:
                continue
            table = parse_xml(self.parts[part])
            maximum = max(maximum, int(table.get("id", "0")))
            names.update(name.casefold() for key in ("name", "displayName") if (name := table.get(key)))
        return maximum, names

    def import_part(self, donor_parts: Mapping[str, bytes], donor_part: str) -> str:
        return self._import_graph(donor_parts, donor_part)[0]

    def _import_graph(self, donor_parts: Mapping[str, bytes], donor_part: str) -> tuple[str, set[str]]:
        source = dict(donor_parts)
        source_types = _types_root(source)
        mapping: dict[str, str] = {}
        table_renames: dict[str, str] = {}
        table_id, table_names = self._table_registry()

        def copy_part(source_part: str) -> str:
            nonlocal table_id
            source_part = _part_name(source_part)
            if source_part in mapping:
                return mapping[source_part]
            if source_part not in source:
                raise ValueError(f"Imported relationship target is missing: {source_part!r}")
            if source_part == "[Content_Types].xml" or source_part.endswith(".rels"):
                raise ValueError(f"Cannot import a manifest as an ordinary part: {source_part!r}")
            target_part = self.allocate(source_part)
            mapping[source_part] = target_part
            payload = source[source_part]
            content_type = _content_type(source_types, source_part)
            if content_type == TABLE_CONTENT_TYPE:
                table = parse_xml(payload)
                old_names = {name for key in ("name", "displayName") if (name := table.get(key))}
                old_name = table.get("displayName") or table.get("name") or "Table"
                new_name = old_name
                index = 0
                while new_name.casefold() in table_names:
                    index += 1
                    suffix = f"_copy{index}"
                    new_name = old_name[: 255 - len(suffix)] + suffix
                table_names.add(new_name.casefold())
                table_id += 1
                table.set("id", str(table_id))
                table.set("name", new_name)
                table.set("displayName", new_name)
                for old in old_names:
                    if old != new_name:
                        table_renames[old.casefold()] = new_name
                payload = serialize_xml(table)
            self.parts[target_part] = payload
            self.ensure_content_type(target_part, content_type)
            relationship_payload = source.get(rels_part(source_part))
            if relationship_payload is None:
                return target_part
            root = parse_xml(relationship_payload)
            if root.tag != f"{{{PACKAGE_REL_NS}}}Relationships":
                raise ValueError(f"Invalid imported relationships for {source_part!r}")
            identifiers: set[str] = set()
            for relationship in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
                identifier = relationship.get("Id", "")
                if not identifier or identifier in identifiers:
                    raise ValueError(f"Duplicate or empty relationship ID in {source_part!r}")
                identifiers.add(identifier)
                mode = relationship.get("TargetMode")
                if mode == "External":
                    continue
                if mode not in {None, "Internal"}:
                    raise ValueError(f"Unsupported relationship target mode: {mode!r}")
                old_target = relationship.get("Target", "")
                destination = copy_part(resolve_target(source_part, old_target))
                fragment = urlsplit(old_target).fragment
                relationship.set(
                    "Target", _relative_target(target_part, destination) + ("#" + fragment if fragment else "")
                )
            self.save_relationships(target_part, root)
            return target_part

        result = copy_part(donor_part)
        if table_renames:
            self._rewrite_table_references(set(mapping.values()), table_renames)
        return result, set(mapping.values())

    def _rewrite_table_references(self, parts: set[str], renames: dict[str, str]) -> None:
        def replace(reference: str) -> str:
            name, separator, rest = reference.partition("[")
            replacement = renames.get(name.casefold())
            return replacement + separator + rest if replacement else reference

        def rewrite(formula: str) -> str:
            if not any(name in formula.casefold() for name in renames):
                return formula
            return _rewrite_range_tokens(formula, replace)

        self._rewrite_graph_formulas(parts, rewrite)

    def _rewrite_graph_formulas(self, parts: set[str], rewrite: Callable[[str], str]) -> None:
        for part in parts:
            if _content_type(self._content_types, part) not in _FORMULA_CONTENT_TYPES:
                continue
            root = parse_xml(self.parts[part])
            changed = False
            for formula in list(root.iter()):
                if formula.tag not in _FORMULA_TAGS or not formula.text:
                    continue
                rewritten = rewrite(formula.text)
                if rewritten != formula.text:
                    formula.text = rewritten
                    _clear_formula_cache(formula)
                    changed = True
            if changed:
                self.parts[part] = serialize_xml(root)

    def clone_part(
        self, part: str, *, source_sheet_name: str | None = None, target_sheet_name: str | None = None
    ) -> str:
        if (source_sheet_name is None) != (target_sheet_name is None):
            raise ValueError("Sheet reference relocation requires both source and target sheet names")
        self.finalize_content_types()
        result, imported = self._import_graph(self.parts, part)
        if source_sheet_name is not None and target_sheet_name is not None:
            self._rewrite_graph_formulas(
                imported, lambda formula: relocate_formula_references(formula, source_sheet_name, target_sheet_name)
            )
        return result

    def remove_part(self, part: str) -> None:
        """Remove only this part and its override; callers own relationship changes."""
        part = _part_name(part)
        if part == "[Content_Types].xml":
            raise ValueError("The package content-types manifest cannot be removed")
        self.parts.pop(part, None)
        for node in list(self._content_types.findall(f"{{{CONTENT_TYPES_NS}}}Override")):
            if unquote(node.get("PartName", "")).lstrip("/") == part:
                self._content_types.remove(node)
                self._content_types_changed = True

    def deduplicate_new_png_media(self, original_parts: set[str]) -> None:
        """Reuse equal PNG payloads, retaining every part that belonged to the original."""
        representatives: dict[bytes, str] = {}
        replacements: dict[str, str] = {}
        candidates = sorted(self.parts, key=lambda part: (part not in original_parts, part))
        for part in candidates:
            if part.endswith(("/", ".rels")) or part == "[Content_Types].xml":
                continue
            if rels_part(part) in self.parts:
                continue
            try:
                content_type = _content_type(self._content_types, part)
            except ValueError:
                continue
            if content_type != "image/png":
                continue
            digest = hashlib.sha256(self.parts[part]).digest()
            representative = representatives.get(digest)
            if representative is None:
                representatives[digest] = part
            elif part not in original_parts and self.parts[representative] == self.parts[part]:
                replacements[part] = representative
        if not replacements:
            return
        for part, payload in tuple(self.parts.items()):
            if not part.endswith(".rels"):
                continue
            owner = _owner_of_rels(part)
            root = parse_xml(payload)
            changed = False
            for relationship in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
                if relationship.get("TargetMode") == "External":
                    continue
                old = relationship.get("Target", "")
                target = resolve_target(owner, old)
                if target in replacements:
                    fragment = urlsplit(old).fragment
                    relationship.set(
                        "Target", _relative_target(owner, replacements[target]) + ("#" + fragment if fragment else "")
                    )
                    changed = True
            if changed:
                self.parts[part] = serialize_xml(root)
        for part in replacements:
            self.remove_part(part)

    def finalize_content_types(self) -> None:
        if self._content_types_changed:
            self.parts["[Content_Types].xml"] = serialize_xml(self._content_types)
