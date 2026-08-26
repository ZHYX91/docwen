"""Word-native list numbering for MD→DOCX conversion.

Creates OOXML numbering definitions (abstractNum + num elements) and
applies numPr (numId + ilvl) to paragraphs, producing proper Word-native
ordered and unordered lists with multi-level support.

Associated findings: F-F1-014, F-F1-015, F-F1-016, F-F1-017, F-F3-023
"""

from __future__ import annotations

import contextlib
import hashlib
import logging
from typing import Any

import lxml.etree as etree
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)

# ── OOXML constants ───────────────────────────────────────────────────────

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"

# Template for an empty numbering.xml
_NUMBERING_XML_TEMPLATE = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/'
    b'wordprocessingml/2006/main"'
    b' xmlns:wpc="http://schemas.microsoft.com/office/word/2010/'
    b'wordprocessingCanvas"'
    b' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
    b'2006/relationships">'
    b"</w:numbering>"
)

# ── Presets ────────────────────────────────────────────────────────────────

BASE_INDENT = 420  # twips (~0.29 inch)
INDENT_INCREMENT = 420  # twips per nesting level

ORDERED_PRESET: dict[str, Any] = {
    "start": "1",
    "numFmt": "decimal",
    "lvlText": "%{level}.",
    "lvlJc": "left",
}

UNORDERED_PRESET: dict[str, Any] = {
    "start": "1",
    "numFmt": "bullet",
    "lvlText": "•",  # Unicode bullet U+2022
    "lvlJc": "left",
}


def _element_ids(elements: list[etree._Element], attribute_name: str) -> set[int]:
    values: set[int] = set()
    attribute = f"{{{WML_NS}}}{attribute_name}"
    for element in elements:
        value = element.get(attribute)
        if value is None:
            continue
        with contextlib.suppress(ValueError):
            values.add(int(value))
    return values


def _numbering_ids_from_root(root: etree._Element) -> tuple[set[int], set[int]]:
    return (
        _element_ids(root.findall(f"{{{WML_NS}}}abstractNum"), "abstractNumId"),
        _element_ids(root.findall(f"{{{WML_NS}}}num"), "numId"),
    )


def _numbering_ids_from_elements(
    abstract_num_elements: list[etree._Element],
    num_elements: list[etree._Element],
) -> tuple[set[int], set[int]]:
    return (
        _element_ids(abstract_num_elements, "abstractNumId"),
        _element_ids(num_elements, "numId"),
    )


def _numbering_ids_from_document(document: Any | None) -> tuple[set[int], set[int]]:
    numbering_root = _numbering_root_from_document(document)
    return _numbering_ids_from_root(numbering_root) if numbering_root is not None else (set(), set())


def _numbering_root_from_document(document: Any | None) -> etree._Element | None:
    if document is None:
        return None
    try:
        return document.part.numbering_part.element
    except (AttributeError, KeyError, NotImplementedError, ValueError):
        return None


def _insert_abstract_num(
    numbering_root: etree._Element,
    abstract_num_element: etree._Element,
) -> None:
    first_num = numbering_root.find(f"{{{WML_NS}}}num")
    if first_num is None:
        numbering_root.append(abstract_num_element)
    else:
        first_num.addprevious(abstract_num_element)


# ═══════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════


class DocxListNumbering:
    """Collects and manages Word-native list numbering definitions.

    During MD→DOCX rendering, list numbering definitions (abstractNum
    and num OOXML elements) are accumulated in memory.  After
    ``Document.save()`` they are written into the DOCX ZIP's
    ``word/numbering.xml`` by :func:`write_numbering_to_docx`.

    This class is **not** a runtime registry; it is a private format
    helper owned by the markdown plugin.
    """

    def __init__(self, document: Any | None = None) -> None:
        # Accumulated OOXML elements — written to ZIP post-save
        self._abstract_num_elements: list[etree._Element] = []
        self._num_elements: list[etree._Element] = []

        # Format tables keyed by level (0‑8)
        self._ordered_formats: dict[int, dict[str, Any]] = {}
        self._unordered_formats: dict[int, dict[str, Any]] = {}

        # Monotonic ID allocators start high but also inspect the actual
        # template. A fixed starting point alone is not collision-safe for
        # user-authored DOCX templates.
        numbering_root = _numbering_root_from_document(document)
        used_abstract_num_ids, used_num_ids = _numbering_ids_from_document(document)
        self._next_abstract_num_id = max({9999, *used_abstract_num_ids}) + 1
        self._next_num_id = max({9999, *used_num_ids}) + 1

        self._build_format_tables(numbering_root)

    # ── Public methods ─────────────────────────────────────────────────

    def create_list_definition(self, level_types: dict[int, str]) -> str:
        """Create a numbering definition for one list group.

        Args:
            level_types:
                ``{level: 'ordered' | 'unordered'}`` for levels 0‑8.
                Missing levels default to ``'ordered'``.

        Returns:
            The allocated ``numId`` as a string, suitable for placement
            in ``w:numId`` on paragraphs belonging to this list group.
        """
        # Assemble 9 level definitions matching the required types
        levels: list[dict[str, Any]] = []
        for lvl in range(9):
            type_ = level_types.get(lvl, "ordered")
            fmt = self._unordered_formats[lvl] if type_ == "unordered" else self._ordered_formats[lvl]
            levels.append(dict(fmt))

        # -- abstractNum --
        abstract_num_id = self._next_abstract_num_id
        self._next_abstract_num_id += 1
        an_elem = _create_abstract_num(abstract_num_id, levels)
        self._abstract_num_elements.append(an_elem)

        # -- num --
        num_id = self._next_num_id
        self._next_num_id += 1
        num_elem = _create_num(num_id, abstract_num_id)
        self._num_elements.append(num_elem)

        logger.debug(
            "Created list def: numId=%s → abstractNumId=%s",
            num_id,
            abstract_num_id,
        )
        return str(num_id)

    @property
    def has_definitions(self) -> bool:
        """``True`` if at least one numbering definition was created."""
        return len(self._num_elements) > 0

    # ── Package-private: used by write_numbering_to_docx ───────────────

    def _get_elements(self) -> tuple[list[etree._Element], list[etree._Element]]:
        """Return copies of ``(abstract_num_elements, num_elements)``."""
        return list(self._abstract_num_elements), list(self._num_elements)

    # ── Internal helpers ───────────────────────────────────────────────

    @staticmethod
    def _extract_level_definition(level_element: etree._Element) -> dict[str, Any]:
        definition: dict[str, Any] = {}
        for child_name, key, default in (
            ("start", "start", "1"),
            ("numFmt", "numFmt", "decimal"),
            ("lvlText", "lvlText", ""),
            ("lvlJc", "lvlJc", "left"),
        ):
            child = level_element.find(f"{{{WML_NS}}}{child_name}")
            if child is not None:
                definition[key] = child.get(f"{{{WML_NS}}}val", default)

        indent = level_element.find(f"{{{WML_NS}}}pPr/{{{WML_NS}}}ind")
        if indent is not None:
            left = indent.get(f"{{{WML_NS}}}left")
            hanging = indent.get(f"{{{WML_NS}}}hanging")
            if left is not None:
                definition["ind_left"] = left
            if hanging is not None:
                definition["ind_hanging"] = hanging

        fonts = level_element.find(f"{{{WML_NS}}}rPr/{{{WML_NS}}}rFonts")
        if fonts is not None:
            ascii_name = fonts.get(f"{{{WML_NS}}}ascii")
            hansi_name = fonts.get(f"{{{WML_NS}}}hAnsi")
            if ascii_name is not None:
                definition["rFonts_ascii"] = ascii_name
            if hansi_name is not None:
                definition["rFonts_hAnsi"] = hansi_name
        return definition

    def _build_format_tables(self, numbering_root: etree._Element | None) -> None:
        """Scan template list formats newest-first, then fill missing levels."""

        if numbering_root is not None:
            abstract_nums: list[tuple[int, etree._Element]] = []
            for abstract_num in numbering_root.findall(f"{{{WML_NS}}}abstractNum"):
                raw_id = abstract_num.get(f"{{{WML_NS}}}abstractNumId")
                if raw_id is None:
                    continue
                with contextlib.suppress(ValueError):
                    abstract_nums.append((int(raw_id), abstract_num))

            for _abstract_num_id, abstract_num in sorted(abstract_nums, reverse=True, key=lambda item: item[0]):
                for level_element in abstract_num.findall(f"{{{WML_NS}}}lvl"):
                    raw_level = level_element.get(f"{{{WML_NS}}}ilvl")
                    if raw_level is None or level_element.get(f"{{{WML_NS}}}tentative") == "1":
                        continue
                    try:
                        level = int(raw_level)
                    except ValueError:
                        continue
                    if not 0 <= level <= 8:
                        continue

                    definition = self._extract_level_definition(level_element)
                    num_fmt = str(definition.get("numFmt", "decimal"))
                    level_text = str(definition.get("lvlText", ""))
                    if ("%" in level_text or num_fmt != "bullet") and level not in self._ordered_formats:
                        self._ordered_formats[level] = definition
                    if (num_fmt == "bullet" or "%" not in level_text) and level not in self._unordered_formats:
                        self._unordered_formats[level] = definition

        self._fill_missing_with_presets()

    def _fill_missing_with_presets(self) -> None:
        """Populate missing format-table levels with deterministic presets."""
        for level in range(9):
            if level not in self._ordered_formats:
                ordered = dict(ORDERED_PRESET)
                ordered["lvlText"] = f"%{level + 1}."
                ordered["ind_left"] = str(BASE_INDENT + INDENT_INCREMENT * level)
                ordered["ind_hanging"] = str(INDENT_INCREMENT)
                self._ordered_formats[level] = ordered

            if level not in self._unordered_formats:
                unordered = dict(UNORDERED_PRESET)
                unordered["ind_left"] = str(BASE_INDENT + INDENT_INCREMENT * level)
                unordered["ind_hanging"] = str(INDENT_INCREMENT)
                self._unordered_formats[level] = unordered


def apply_list_to_paragraph(paragraph, num_id: str, level: int) -> None:
    """Apply Word-native list numbering to *paragraph*.

    Writes ``w:numPr`` → ``w:numId`` + ``w:ilvl`` into the paragraph's
    ``w:pPr`` element so that Word activates the numbering definition
    referenced by *num_id* at the given *level* (0‑based).

    This is a standalone helper usable from any rendering code.
    """
    p_elem = paragraph._p

    # Get or create w:pPr
    pPr = p_elem.find(f"{{{WML_NS}}}pPr")
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)

    # Remove existing numPr if present
    existing = pPr.find(f"{{{WML_NS}}}numPr")
    if existing is not None:
        pPr.remove(existing)

    # Build w:numPr
    numPr = OxmlElement("w:numPr")

    ilvl_elem = OxmlElement("w:ilvl")
    ilvl_elem.set(qn("w:val"), str(level))
    numPr.append(ilvl_elem)

    numId_elem = OxmlElement("w:numId")
    numId_elem.set(qn("w:val"), num_id)
    numPr.append(numId_elem)

    pPr.insert(0, numPr)


# ═══════════════════════════════════════════════════════════════════════════
# Post-save writeback
# ═══════════════════════════════════════════════════════════════════════════


def write_numbering_to_docx(
    docx_path: str,
    numbering: DocxListNumbering,
) -> None:
    """Write accumulated numbering definitions into a saved DOCX file.

    Opens *docx_path* as a ZIP, ensures ``word/numbering.xml`` exists,
    appends the ``abstractNum`` and ``num`` elements collected by
    *numbering* during rendering, updates relationships and content types,
    then writes back.

    Must be called **after** ``Document.save()`` so the ZIP contains
    all body content.
    """
    import tempfile
    import zipfile
    from pathlib import Path

    if not numbering.has_definitions:
        return

    abstract_elems, num_elems = numbering._get_elements()
    original = Path(docx_path)

    with tempfile.TemporaryDirectory(prefix=".dw-numbering-", dir=original.parent) as tmpdir:
        tmp_path = Path(tmpdir) / "numbering_writeback.docx"

        with (
            zipfile.ZipFile(str(original), "r") as zf_in,
            zipfile.ZipFile(str(tmp_path), "w", zipfile.ZIP_DEFLATED) as zf_out,
        ):
            names = set(zf_in.namelist())

            # Read relationships and content types once
            rels_raw = zf_in.read("word/_rels/document.xml.rels")
            ct_raw = zf_in.read("[Content_Types].xml")
            rels_root = etree.fromstring(rels_raw)
            ct_root = etree.fromstring(ct_raw)

            has_numbering = "word/numbering.xml" in names

            if has_numbering:
                raw = zf_in.read("word/numbering.xml")
                num_root = etree.fromstring(raw)
            else:
                num_root = etree.fromstring(_NUMBERING_XML_TEMPLATE)

            existing_abstract_num_ids, existing_num_ids = _numbering_ids_from_root(num_root)
            generated_abstract_num_ids, generated_num_ids = _numbering_ids_from_elements(
                abstract_elems,
                num_elems,
            )
            abstract_collisions = existing_abstract_num_ids & generated_abstract_num_ids
            num_collisions = existing_num_ids & generated_num_ids
            if abstract_collisions or num_collisions:
                raise ValueError(
                    "Generated list numbering IDs collide with the saved template: "
                    f"abstractNumId={sorted(abstract_collisions)}, numId={sorted(num_collisions)}"
                )

            # CT_Numbering schema order is numPicBullet*, abstractNum*, num*.
            for elem in abstract_elems:
                _insert_abstract_num(num_root, elem)
            for elem in num_elems:
                num_root.append(elem)

            num_bytes = etree.tostring(
                num_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

            if not has_numbering:
                _ensure_relationship(rels_root, "numbering.xml", NUMBERING_REL_TYPE)
                _ensure_content_type(ct_root, "/word/numbering.xml", NUMBERING_CONTENT_TYPE)

            modified_rels = etree.tostring(
                rels_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            modified_ct = etree.tostring(
                ct_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

            # Rebuild ZIP
            for item in zf_in.infolist():
                if item.filename == "word/numbering.xml":
                    zf_out.writestr(item, num_bytes)
                elif item.filename == "word/_rels/document.xml.rels":
                    zf_out.writestr(item, modified_rels)
                elif item.filename == "[Content_Types].xml":
                    zf_out.writestr(item, modified_ct)
                else:
                    zf_out.writestr(item, zf_in.read(item.filename))

            if not has_numbering:
                zf_out.writestr("word/numbering.xml", num_bytes)

        tmp_path.replace(original)


# ═══════════════════════════════════════════════════════════════════════════
# OOXML element factories  (private)
# ═══════════════════════════════════════════════════════════════════════════


def _create_abstract_num(
    abstract_num_id: int,
    levels: list[dict[str, Any]],
) -> etree._Element:
    """Create a ``w:abstractNum`` element with 9 ``w:lvl`` children."""
    an = etree.Element(f"{{{WML_NS}}}abstractNum")
    an.set(f"{{{WML_NS}}}abstractNumId", str(abstract_num_id))

    # nsid — required by OOXML (unique per abstractNum)
    nsid = etree.SubElement(an, f"{{{WML_NS}}}nsid")
    nsid_value = hashlib.sha256(f"docwen-list:{abstract_num_id}".encode()).hexdigest()[:8].upper()
    nsid.set(f"{{{WML_NS}}}val", nsid_value)

    # multiLevelType
    mlt = etree.SubElement(an, f"{{{WML_NS}}}multiLevelType")
    mlt.set(f"{{{WML_NS}}}val", "hybridMultilevel")

    for lvl_idx, lvl_def in enumerate(levels):
        an.append(_create_lvl(lvl_idx, lvl_def))

    return an


def _create_lvl(level: int, lvl_def: dict[str, Any]) -> etree._Element:
    """Create a single ``w:lvl`` element from *lvl_def*."""
    lvl = etree.Element(f"{{{WML_NS}}}lvl")
    lvl.set(f"{{{WML_NS}}}ilvl", str(level))

    e = etree.SubElement(lvl, f"{{{WML_NS}}}start")
    e.set(f"{{{WML_NS}}}val", lvl_def.get("start", "1"))

    e = etree.SubElement(lvl, f"{{{WML_NS}}}numFmt")
    e.set(f"{{{WML_NS}}}val", lvl_def.get("numFmt", "decimal"))

    e = etree.SubElement(lvl, f"{{{WML_NS}}}lvlText")
    e.set(f"{{{WML_NS}}}val", lvl_def.get("lvlText", f"%{level + 1}."))

    e = etree.SubElement(lvl, f"{{{WML_NS}}}lvlJc")
    e.set(f"{{{WML_NS}}}val", lvl_def.get("lvlJc", "left"))

    # pPr → ind
    pPr = etree.SubElement(lvl, f"{{{WML_NS}}}pPr")
    ind = etree.SubElement(pPr, f"{{{WML_NS}}}ind")
    ind.set(
        f"{{{WML_NS}}}left",
        lvl_def.get("ind_left", str(BASE_INDENT + INDENT_INCREMENT * level)),
    )
    ind.set(
        f"{{{WML_NS}}}hanging",
        lvl_def.get("ind_hanging", str(INDENT_INCREMENT)),
    )

    if lvl_def.get("numFmt") == "bullet":
        ascii_name = lvl_def.get("rFonts_ascii")
        hansi_name = lvl_def.get("rFonts_hAnsi")
        if ascii_name or hansi_name:
            r_pr = etree.SubElement(lvl, f"{{{WML_NS}}}rPr")
            r_fonts = etree.SubElement(r_pr, f"{{{WML_NS}}}rFonts")
            r_fonts.set(f"{{{WML_NS}}}ascii", str(ascii_name or hansi_name))
            r_fonts.set(f"{{{WML_NS}}}hAnsi", str(hansi_name or ascii_name))
            r_fonts.set(f"{{{WML_NS}}}hint", "default")

    return lvl


def _create_num(num_id: int, abstract_num_id: int) -> etree._Element:
    """Create a ``w:num`` element referencing an ``abstractNum``."""
    num = etree.Element(f"{{{WML_NS}}}num")
    num.set(f"{{{WML_NS}}}numId", str(num_id))

    an_ref = etree.SubElement(num, f"{{{WML_NS}}}abstractNumId")
    an_ref.set(f"{{{WML_NS}}}val", str(abstract_num_id))

    return num


# ═══════════════════════════════════════════════════════════════════════════
# Relationship / content-type helpers  (private)
# ═══════════════════════════════════════════════════════════════════════════


def _ensure_relationship(
    rels_root: etree._Element,
    target: str,
    rel_type: str,
) -> None:
    """Add a ``<Relationship>`` for *target* if none exists."""
    for rel in rels_root.findall(f"{{{_RELS_NS}}}Relationship"):
        if rel.get("Target") == target:
            return
    max_id = 0
    for rel in rels_root.findall(f"{{{_RELS_NS}}}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            with contextlib.suppress(ValueError):
                max_id = max(max_id, int(rid[3:]))
    elem = etree.SubElement(rels_root, f"{{{_RELS_NS}}}Relationship")
    elem.set("Id", f"rId{max_id + 1}")
    elem.set("Type", rel_type)
    elem.set("Target", target)


def _ensure_content_type(
    ct_root: etree._Element,
    part_name: str,
    content_type: str,
) -> None:
    """Add an ``<Override>`` for *part_name* if none exists."""
    for override in ct_root.findall(f"{{{_CT_NS}}}Override"):
        if override.get("PartName") == part_name:
            return
    elem = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
    elem.set("PartName", part_name)
    elem.set("ContentType", content_type)
