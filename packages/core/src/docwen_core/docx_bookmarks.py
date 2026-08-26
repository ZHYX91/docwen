"""Safe physical bookmark inventory and proof for WordprocessingML packages."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Literal

BOOKMARK_ID_MAX = 2_147_483_647
BOOKMARK_NAME_MAX_LENGTH = 40

type BookmarkIdKey = tuple[Literal["numeric", "text"], int | str]

_BOOKMARK_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,39}$")
_DECIMAL_ID_RE = re.compile(r"^[+-]?[0-9]+$")


@dataclass(frozen=True, slots=True)
class DocxBookmarkOccurrence:
    """One physical bookmark start or end in a specific package part."""

    part_name: str
    position: int
    element: Any
    raw_id: str
    id_key: BookmarkIdKey
    name: str | None = None


@dataclass(frozen=True, slots=True)
class DocxBookmarkInventory:
    """Package-wide bookmark occurrences, including orphan starts and ends."""

    starts: tuple[DocxBookmarkOccurrence, ...]
    ends: tuple[DocxBookmarkOccurrence, ...]

    @property
    def used_id_keys(self) -> frozenset[BookmarkIdKey]:
        return frozenset(item.id_key for item in (*self.starts, *self.ends))

    @property
    def used_name_keys(self) -> frozenset[str]:
        return frozenset(item.name.casefold() for item in self.starts if item.name is not None)

    def starts_named(self, name: str) -> tuple[DocxBookmarkOccurrence, ...]:
        name_key = name.casefold()
        return tuple(item for item in self.starts if item.name is not None and item.name.casefold() == name_key)

    def starts_with_id(self, id_key: BookmarkIdKey) -> tuple[DocxBookmarkOccurrence, ...]:
        return tuple(item for item in self.starts if item.id_key == id_key)

    def ends_with_id(self, id_key: BookmarkIdKey) -> tuple[DocxBookmarkOccurrence, ...]:
        return tuple(item for item in self.ends if item.id_key == id_key)


@dataclass(frozen=True, slots=True)
class DocxBookmarkProof:
    """A globally unique, balanced, ordered bookmark range proof."""

    valid: bool
    start: DocxBookmarkOccurrence | None = None
    end: DocxBookmarkOccurrence | None = None


def bookmark_id_key(raw_id: str) -> BookmarkIdKey:
    """Canonicalize decimal lexical IDs while preserving other strings exactly."""

    stripped = raw_id.strip()
    if _DECIMAL_ID_RE.fullmatch(stripped) is not None:
        return ("numeric", int(stripped, 10))
    return ("text", raw_id)


def is_legal_bookmark_name(name: str) -> bool:
    """Return whether a semantic bookmark name fits the portable Word subset."""

    return len(name) <= BOOKMARK_NAME_MAX_LENGTH and _BOOKMARK_NAME_RE.fullmatch(name) is not None


def prove_bookmark_name(
    inventory: DocxBookmarkInventory,
    name: str,
    *,
    scope_element: Any | None = None,
) -> DocxBookmarkProof:
    """Prove one bookmark name has one globally unambiguous physical range."""

    if not is_legal_bookmark_name(name):
        return DocxBookmarkProof(False)
    named_starts = inventory.starts_named(name)
    if len(named_starts) != 1:
        return DocxBookmarkProof(False)
    start = named_starts[0]
    if start.id_key[0] != "numeric":
        return DocxBookmarkProof(False, start=start)
    numeric_id = start.id_key[1]
    if not isinstance(numeric_id, int) or not 0 <= numeric_id <= BOOKMARK_ID_MAX:
        return DocxBookmarkProof(False, start=start)
    matching_starts = inventory.starts_with_id(start.id_key)
    matching_ends = inventory.ends_with_id(start.id_key)
    if len(matching_starts) != 1 or len(matching_ends) != 1:
        return DocxBookmarkProof(False, start=start)
    end = matching_ends[0]
    if start.part_name != end.part_name or end.position <= start.position:
        return DocxBookmarkProof(False, start=start, end=end)
    if scope_element is not None:
        scoped_element_ids = {id(element) for element in scope_element.iter()}
        if id(start.element) not in scoped_element_ids or id(end.element) not in scoped_element_ids:
            return DocxBookmarkProof(False, start=start, end=end)
    return DocxBookmarkProof(True, start=start, end=end)


def build_docx_bookmark_inventory(document: Any) -> DocxBookmarkInventory:
    """Inventory bookmarks in every reachable XML part below ``/word``.

    Generic OPC parts such as footnotes, endnotes, and comments expose only a
    byte blob.  Only parts whose path and content type both identify XML are
    parsed; binary and non-Word package content is never inspected.
    """

    roots: dict[str, Any] = {str(document.part.partname): document.element}
    for part in document.part.package.parts:
        part_name = str(part.partname)
        content_type = str(part.content_type).lower()
        if (
            not part_name.lower().startswith("/word/")
            or not part_name.lower().endswith(".xml")
            or not content_type.endswith("xml")
        ):
            continue
        if part_name in roots:
            continue
        element = getattr(part, "element", None)
        if element is None:
            element = getattr(part, "_element", None)
        if element is None:
            element = _parse_xml_part(part, part_name=part_name)
        roots[part_name] = element
    return _inventory_from_roots(tuple(roots.items()))


def _parse_xml_part(part: Any, *, part_name: str) -> Any:
    from lxml import etree

    try:
        blob = part.blob
        if not isinstance(blob, bytes) or not blob:
            raise ValueError
        parser = etree.XMLParser(
            load_dtd=False,
            no_network=True,
            recover=False,
            resolve_entities=False,
        )
        return etree.fromstring(blob, parser=parser)
    except (AttributeError, TypeError, ValueError, etree.XMLSyntaxError):
        raise ValueError(f"Cannot inventory WordprocessingML XML part {part_name}.") from None


def _inventory_from_roots(roots: tuple[tuple[str, Any], ...]) -> DocxBookmarkInventory:
    from docx.oxml.ns import qn

    start_tag = qn("w:bookmarkStart")
    end_tag = qn("w:bookmarkEnd")
    id_attribute = qn("w:id")
    name_attribute = qn("w:name")
    starts: list[DocxBookmarkOccurrence] = []
    ends: list[DocxBookmarkOccurrence] = []
    for part_name, root in roots:
        for position, element in enumerate(root.iter()):
            if element.tag not in {start_tag, end_tag}:
                continue
            raw_id = element.get(id_attribute) or ""
            occurrence = DocxBookmarkOccurrence(
                part_name=part_name,
                position=position,
                element=element,
                raw_id=raw_id,
                id_key=bookmark_id_key(raw_id),
                name=(element.get(name_attribute) if element.tag == start_tag else None),
            )
            if element.tag == start_tag:
                starts.append(occurrence)
            else:
                ends.append(occurrence)
    return DocxBookmarkInventory(starts=tuple(starts), ends=tuple(ends))


__all__ = [
    "BOOKMARK_ID_MAX",
    "BOOKMARK_NAME_MAX_LENGTH",
    "BookmarkIdKey",
    "DocxBookmarkInventory",
    "DocxBookmarkOccurrence",
    "DocxBookmarkProof",
    "bookmark_id_key",
    "build_docx_bookmark_inventory",
    "is_legal_bookmark_name",
    "prove_bookmark_name",
]
