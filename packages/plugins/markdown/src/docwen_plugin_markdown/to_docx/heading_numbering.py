"""Apply native numbering only to headings created by this conversion.

Template Heading styles remain shared, immutable identities. Request-owned
paragraph bindings survive save without adding private markers to the file.
"""

from __future__ import annotations

import contextlib
import hashlib
import logging
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

import lxml.etree as etree
from docx.document import Document
from docx.text.paragraph import Paragraph

from docwen_core.docx_numbering_parts import open_numbering_part
from docwen_core.docx_numbering_structure import append_numbering_definitions
from docwen_core.docx_package_parts import DocxPackageParts
from docwen_core.file_io import replace_authorized_file
from docwen_core.tempfiles import TemporaryDirectory
from docwen_core.text.numbering_word_adapter import (
    TranslationResult,
    WordNumberingLevel,
)
from docwen_plugin_markdown.to_docx.direction import paragraph_is_rtl
from docwen_plugin_markdown.to_docx.numbering import WML_NS

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════════
# Constants
# ═══════════════════════════════════════════════════════════════════════════

BASE_INDENT = 420  # twips
INDENT_INCREMENT = 420  # twips per level


@dataclass(frozen=True)
class _HeadingTarget:
    path: tuple[int, ...]
    level: int
    attributes: tuple[tuple[str, str], ...]
    properties: bytes
    right_to_left: bool


@dataclass(frozen=True)
class HeadingNumberingTargets:
    """A save-boundary snapshot of renderer-owned heading paragraphs."""

    paragraph_count: int
    headings: tuple[_HeadingTarget, ...]


def _paragraph_properties(paragraph: etree._Element) -> bytes:
    properties = paragraph.find(f"{{{WML_NS}}}pPr")
    if properties is None:
        return b""
    return etree.tostring(properties, method="c14n", exclusive=True)


def _paragraph_attributes(paragraph: etree._Element) -> tuple[tuple[str, str], ...]:
    return tuple(sorted((str(name), str(value)) for name, value in paragraph.attrib.items()))


def capture_heading_numbering_targets(
    document: Document,
    headings: Iterable[tuple[Paragraph, int]],
) -> HeadingNumberingTargets:
    """Capture exact paragraphs after template filling and before save.

    All recorded paragraphs must have been moved into the final main story.
    Runs are deliberately not fingerprinted: note writeback can replace a
    repeated note reference with its equivalent NOTEREF field before numbering.
    Paragraph identity, structure and properties must remain unchanged.
    """
    root = document.element
    targets: list[_HeadingTarget] = []
    seen: set[tuple[int, ...]] = set()
    for paragraph, level in headings:
        if not 1 <= level <= 9:
            raise ValueError(f"Invalid generated heading level: {level}.")
        element = paragraph._p
        path: list[int] = []
        ancestor: etree._Element = element
        while ancestor is not root:
            parent = ancestor.getparent()
            if parent is None:
                raise ValueError("Generated heading is not attached to the final document.")
            path.append(parent.index(ancestor))
            ancestor = parent
        identity = tuple(reversed(path))
        if identity in seen:
            raise ValueError("Generated heading was recorded more than once.")
        seen.add(identity)
        targets.append(
            _HeadingTarget(
                identity,
                level,
                _paragraph_attributes(element),
                _paragraph_properties(element),
                paragraph_is_rtl(paragraph),
            )
        )
    return HeadingNumberingTargets(sum(1 for _ in root.iter(f"{{{WML_NS}}}p")), tuple(targets))


# ═══════════════════════════════════════════════════════════════════════════
# OOXML element factories
# ═══════════════════════════════════════════════════════════════════════════


def create_heading_abstract_num(
    abstract_num_id: int,
    translation_result: TranslationResult,
    *,
    right_to_left_levels: frozenset[int] = frozenset(),
) -> etree._Element:
    """Create a ``w:abstractNum`` element from the translation result's levels.

    Levels deliberately have no ``w:pStyle`` link: the same Heading style
    can be used by fixed template content outside this conversion's scope.
    Generated headings select their level through direct paragraph numPr.

    Args:
        abstract_num_id: The allocated abstractNumId to use.
        translation_result: Heading numbering levels from the translator.
    """
    an = etree.Element(f"{{{WML_NS}}}abstractNum")
    an.set(f"{{{WML_NS}}}abstractNumId", str(abstract_num_id))

    # nsid — required by OOXML (unique per abstractNum)
    nsid = etree.SubElement(an, f"{{{WML_NS}}}nsid")
    nsid_val = hashlib.md5(f"heading_{abstract_num_id}".encode()).hexdigest()[:8].upper()
    nsid.set(f"{{{WML_NS}}}val", nsid_val)

    # A single multilevel sequence shared by generated headings.
    mlt = etree.SubElement(an, f"{{{WML_NS}}}multiLevelType")
    mlt.set(f"{{{WML_NS}}}val", "hybridMultilevel")

    for level in translation_result.levels:
        an.append(_create_heading_lvl(level, right_to_left=level.ilvl in right_to_left_levels))

    return an


def _create_heading_lvl(level: WordNumberingLevel, *, right_to_left: bool) -> etree._Element:
    """Create a single ``w:lvl`` element from a ``WordNumberingLevel``."""
    lvl = etree.Element(f"{{{WML_NS}}}lvl")
    lvl.set(f"{{{WML_NS}}}ilvl", str(level.ilvl))

    # start
    e = etree.SubElement(lvl, f"{{{WML_NS}}}start")
    e.set(f"{{{WML_NS}}}val", level.start)

    # numFmt
    e = etree.SubElement(lvl, f"{{{WML_NS}}}numFmt")
    e.set(f"{{{WML_NS}}}val", level.num_fmt)

    # suff — what follows the number
    e = etree.SubElement(lvl, f"{{{WML_NS}}}suff")
    e.set(f"{{{WML_NS}}}val", level.suff)

    # lvlText — the format template (e.g. "%1、")
    e = etree.SubElement(lvl, f"{{{WML_NS}}}lvlText")
    e.set(f"{{{WML_NS}}}val", level.lvl_text)

    # lvlJc
    e = etree.SubElement(lvl, f"{{{WML_NS}}}lvlJc")
    e.set(f"{{{WML_NS}}}val", "right" if right_to_left else "left")

    # pPr → ind (hanging indent for clean wrapping)
    pPr = etree.SubElement(lvl, f"{{{WML_NS}}}pPr")
    ind = etree.SubElement(pPr, f"{{{WML_NS}}}ind")
    left_val = BASE_INDENT + INDENT_INCREMENT * level.ilvl
    ind.set(f"{{{WML_NS}}}{'right' if right_to_left else 'left'}", str(left_val))
    ind.set(f"{{{WML_NS}}}hanging", str(INDENT_INCREMENT))

    return lvl


def create_heading_num(num_id: int, abstract_num_id: int) -> etree._Element:
    """Create a ``w:num`` element referencing an ``abstractNum``."""
    num = etree.Element(f"{{{WML_NS}}}num")
    num.set(f"{{{WML_NS}}}numId", str(num_id))

    an_ref = etree.SubElement(num, f"{{{WML_NS}}}abstractNumId")
    an_ref.set(f"{{{WML_NS}}}val", str(abstract_num_id))

    return num


# ═══════════════════════════════════════════════════════════════════════════
# Request-owned paragraph numbering
# ═══════════════════════════════════════════════════════════════════════════

# CT_PPrBase child order — numPr comes after these elements:
_PPR_BEFORE_NUMPR = {
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
}


def _insert_numpr(pPr: etree._Element, numPr: etree._Element) -> None:
    """Insert numPr at the correct schema position within pPr.

    CT_PPrBase requires numPr after pStyle/keepNext/etc but before
    ind/spacing/jc.  Find the first child that sorts after numPr and
    insert before it; if none, append.
    """
    for i, child in enumerate(pPr):
        tag = etree.QName(child).localname
        if tag not in _PPR_BEFORE_NUMPR:
            pPr.insert(i, numPr)
            return
    pPr.append(numPr)


def _resolve_targets(
    document_root: etree._Element,
    targets: HeadingNumberingTargets,
) -> list[tuple[etree._Element, int]]:
    if sum(1 for _ in document_root.iter(f"{{{WML_NS}}}p")) != targets.paragraph_count:
        raise ValueError("Document paragraphs changed after heading numbering targets were captured.")
    resolved: list[tuple[etree._Element, int]] = []
    for target in targets.headings:
        element = document_root
        for index in target.path:
            if not 0 <= index < len(element):
                raise ValueError("Generated heading no longer exists at its captured position.")
            element = element[index]
        if (
            element.tag != f"{{{WML_NS}}}p"
            or _paragraph_attributes(element) != target.attributes
            or _paragraph_properties(element) != target.properties
        ):
            raise ValueError("Generated heading identity or properties changed before numbering writeback.")
        resolved.append((element, target.level - 1))
    return resolved


def _apply_heading_number(paragraph: etree._Element, *, num_id: int, level: int | None) -> None:
    properties = paragraph.find(f"{{{WML_NS}}}pPr")
    if properties is None:
        properties = etree.Element(f"{{{WML_NS}}}pPr")
        paragraph.insert(0, properties)
    for existing in properties.findall(f"{{{WML_NS}}}numPr"):
        properties.remove(existing)
    numbering = etree.Element(f"{{{WML_NS}}}numPr")
    if level is not None:
        etree.SubElement(numbering, f"{{{WML_NS}}}ilvl").set(f"{{{WML_NS}}}val", str(level))
    etree.SubElement(numbering, f"{{{WML_NS}}}numId").set(f"{{{WML_NS}}}val", str(num_id))
    _insert_numpr(properties, numbering)


# ═══════════════════════════════════════════════════════════════════════════
# Helpers — numbering.xml element ordering
# ═══════════════════════════════════════════════════════════════════════════


def _allocate_ids(num_root: etree._Element, reference_roots: Iterable[etree._Element]) -> tuple[int, int]:
    """Find free abstractNumId and numId in an existing numbering root.

    Scans all existing ``abstractNumId`` and ``numId`` attribute values,
    then picks the next available integer starting at 20 000 (above the
    list-numbering range which starts at 10 000).

    Returns:
        ``(abstract_num_id, num_id)`` — typically the same number, since
        Word allows abstractNumId and numId to share values.
    """
    used: set[int] = set()
    for element_name, attribute in (("abstractNum", "abstractNumId"), ("num", "numId")):
        for element in num_root.findall(f"{{{WML_NS}}}{element_name}"):
            val = element.get(f"{{{WML_NS}}}{attribute}")
            if val is not None:
                with contextlib.suppress(ValueError):
                    used.add(int(val))
    # Dangling template references must stay dangling, rather than silently
    # acquire the new request's numbering when an ID happens to be reused.
    for root in (num_root, *reference_roots):
        for tag in ("abstractNumId", "numId"):
            for element in root.iter(f"{{{WML_NS}}}{tag}"):
                val = element.get(f"{{{WML_NS}}}val")
                if val is not None:
                    with contextlib.suppress(ValueError):
                        used.add(int(val))
    candidate = 20000
    while candidate in used:
        candidate += 1
    return candidate, candidate


# ═══════════════════════════════════════════════════════════════════════════
# Post-save ZIP writeback
# ═══════════════════════════════════════════════════════════════════════════


def _parse_xml(raw: bytes) -> etree._Element:
    return etree.fromstring(raw, etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False))


def _xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def write_heading_numbering_to_docx(
    docx_path: str,
    translation_result: TranslationResult,
    *,
    targets: HeadingNumberingTargets,
) -> None:
    """Write native numbering to the exact generated headings captured before save.

    Run after list-numbering writeback so IDs are allocated against the final
    package inventory. Shared template styles and all unrelated parts are
    copied unchanged. Target mismatch fails before replacing the original.
    """
    supported_levels = {level.ilvl for level in translation_result.levels}
    if not targets.headings:
        return
    has_supported_targets = any(target.level - 1 in supported_levels for target in targets.headings)

    original = Path(docx_path)
    with zipfile.ZipFile(original) as archive:
        infos = archive.infolist()
        if len({item.filename for item in infos}) != len(infos):
            raise ValueError("DOCX contains duplicate ZIP members.")
        data = {item.filename: archive.read(item.filename) for item in infos}
    document_part = DocxPackageParts(data).document()
    document_root = document_part.element
    resolved = _resolve_targets(document_root, targets)
    part = open_numbering_part(data, create=has_supported_targets)
    num_root = part.element if part is not None else etree.Element(f"{{{WML_NS}}}numbering")

    num_id = 0
    if has_supported_targets:
        reference_roots: list[etree._Element] = []
        for name, raw in data.items():
            if name.endswith(".xml") and (part is None or name != part.name):
                # Opaque, malformed optional parts are outside this writer's scope.
                # Their bytes are retained; valid Word parts reserve dangling IDs.
                with contextlib.suppress(etree.XMLSyntaxError):
                    reference_roots.append(_parse_xml(raw))
        abstract_num_id, num_id = _allocate_ids(num_root, reference_roots)
        directions: dict[int, bool] = {}
        for target in targets.headings:
            level = target.level - 1
            previous = directions.setdefault(level, target.right_to_left)
            if level in supported_levels and previous != target.right_to_left:
                raise ValueError("Generated headings at one numbering level have conflicting text directions.")
        abstract_num = create_heading_abstract_num(
            abstract_num_id,
            translation_result,
            right_to_left_levels=frozenset(level for level, rtl in directions.items() if rtl),
        )
        append_numbering_definitions(num_root, (abstract_num,), (create_heading_num(num_id, abstract_num_id),))
        assert part is not None
        data[part.name] = _xml_bytes(num_root)
    for paragraph, level in resolved:
        if level in supported_levels:
            _apply_heading_number(paragraph, num_id=num_id, level=level)
        else:
            # A scheme's omitted/unsupported level must not inherit numbering
            # from a shared template Heading style. numId=0 cancels inheritance.
            _apply_heading_number(paragraph, num_id=0, level=None)

    data[document_part.name] = _xml_bytes(document_root)

    with TemporaryDirectory(prefix=".dw-heading-numbering-", dir=original.parent) as tmpdir:
        tmp_path = Path(tmpdir) / "heading_numbering_writeback.docx"
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for info in infos:
                archive.writestr(info, data[info.filename])
            existing = {info.filename for info in infos}
            for name in data.keys() - existing:
                archive.writestr(name, data[name])
        replace_authorized_file(tmp_path, original)
    logger.info("Numbered %d generated headings with numId=%s", len(resolved), num_id)
