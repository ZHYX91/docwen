from __future__ import annotations

from collections.abc import Iterable
from pathlib import PurePosixPath
from typing import NoReturn
from xml.etree import ElementTree

RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL = f"{{{RELS}}}"
_CT = f"{{{CONTENT_TYPES}}}"
_NUMBERING_REL = f"{OFFICE_RELS}/numbering"
_STYLES_REL = f"{OFFICE_RELS}/styles"
_DOCUMENT_REL = f"{OFFICE_RELS}/officeDocument"
_NUMBERING_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
_STYLES_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_DOCUMENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"


class V4OoxmlOpcError(ValueError):
    """An OPC content-type or relationship graph is incomplete or ambiguous."""


def _fail(code: str) -> NoReturn:
    raise V4OoxmlOpcError(code)


def _relationship_index(root: ElementTree.Element, *, base: str, names: set[str]) -> list[dict[str, str]]:
    if root.tag != f"{_REL}Relationships":
        _fail("ooxml_opc_relationship_root_invalid")
    result: list[dict[str, str]] = []
    ids: set[str] = set()
    for item in root:
        if item.tag != f"{_REL}Relationship":
            _fail("ooxml_opc_relationship_child_invalid")
        relationship_id = item.get("Id")
        relationship_type = item.get("Type")
        target = item.get("Target")
        target_mode = item.get("TargetMode")
        if (
            not relationship_id
            or relationship_id in ids
            or not relationship_type
            or not target
            or target_mode not in {None, "External"}
        ):
            _fail("ooxml_opc_relationship_invalid")
        ids.add(relationship_id)
        if target_mode is None:
            path = PurePosixPath(target)
            if path.is_absolute() or "\\" in target or ":" in target or "#" in target or "?" in target:
                _fail("ooxml_opc_relationship_target_invalid")
            resolved_parts = [part for part in PurePosixPath(base).parts if part]
            for part in path.parts:
                if part in {"", "."}:
                    _fail("ooxml_opc_relationship_target_invalid")
                if part == "..":
                    if not resolved_parts:
                        _fail("ooxml_opc_relationship_target_invalid")
                    resolved_parts.pop()
                else:
                    resolved_parts.append(part)
            resolved = "/".join(resolved_parts)
            if resolved not in names:
                _fail("ooxml_opc_relationship_target_missing")
        result.append({"type": relationship_type, "target": target, "mode": target_mode or "Internal"})
    return result


def _single_relationship(
    relationships: Iterable[dict[str, str]],
    *,
    relationship_type: str,
    target: str,
) -> bool:
    matches = [item for item in relationships if item["type"] == relationship_type or item["target"] == target]
    return len(matches) == 1 and matches[0] == {"type": relationship_type, "target": target, "mode": "Internal"}


def prove_opc(
    names: set[str],
    content_types: ElementTree.Element,
    root_relationships: ElementTree.Element,
    document_relationships: ElementTree.Element,
    *,
    numbering_present: bool,
) -> None:
    if content_types.tag != f"{_CT}Types":
        _fail("ooxml_opc_content_types_root_invalid")
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for item in content_types:
        if item.tag == f"{_CT}Default":
            extension, content_type = item.get("Extension"), item.get("ContentType")
            if not extension or not content_type or extension.casefold() in defaults:
                _fail("ooxml_opc_default_content_type_invalid")
            defaults[extension.casefold()] = content_type
        elif item.tag == f"{_CT}Override":
            part_name, content_type = item.get("PartName"), item.get("ContentType")
            if not part_name or not part_name.startswith("/") or not content_type or part_name in overrides:
                _fail("ooxml_opc_override_content_type_invalid")
            overrides[part_name] = content_type
        else:
            _fail("ooxml_opc_content_type_child_invalid")
    for name in names - {"[Content_Types].xml"}:
        extension = name.rsplit(".", 1)[-1].casefold() if "." in PurePosixPath(name).name else ""
        if f"/{name}" not in overrides and extension not in defaults:
            _fail("ooxml_opc_part_content_type_missing")
    if overrides.get("/word/document.xml") != _DOCUMENT_TYPE:
        _fail("ooxml_document_content_type_invalid")
    if overrides.get("/word/styles.xml") != _STYLES_TYPE:
        _fail("ooxml_styles_content_type_invalid")
    root = _relationship_index(root_relationships, base="", names=names)
    document = _relationship_index(document_relationships, base="word", names=names)
    if not _single_relationship(root, relationship_type=_DOCUMENT_REL, target="word/document.xml"):
        _fail("ooxml_document_root_relationship_invalid")
    if not _single_relationship(document, relationship_type=_STYLES_REL, target="styles.xml"):
        _fail("ooxml_styles_relationship_invalid")
    numbering_type = overrides.get("/word/numbering.xml")
    numbering_candidates = [
        item for item in document if item["type"] == _NUMBERING_REL or item["target"] == "numbering.xml"
    ]
    numbering_ok = _single_relationship(document, relationship_type=_NUMBERING_REL, target="numbering.xml")
    if numbering_present:
        if "word/numbering.xml" not in names or numbering_type != _NUMBERING_TYPE or not numbering_ok:
            _fail("ooxml_numbering_opc_support_invalid")
    elif "word/numbering.xml" in names or numbering_type is not None or numbering_candidates:
        _fail("ooxml_disabled_numbering_part_present")


__all__ = ["V4OoxmlOpcError", "prove_opc"]
