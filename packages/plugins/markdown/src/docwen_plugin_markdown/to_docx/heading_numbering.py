"""Word-native heading numbering injection for MD→DOCX conversion.

Creates OOXML numbering definitions with pStyle-linked levels for heading
numbering, and injects numPr into Heading style definitions in styles.xml.
Called after Document.save() to perform post-save ZIP writeback.
"""

from __future__ import annotations

import contextlib
import hashlib
import logging
import tempfile
import zipfile
from collections.abc import Mapping
from dataclasses import replace
from pathlib import Path

import lxml.etree as etree

from docwen_core.text.numbering_word_adapter import (
    TranslationResult,
    WordNumberingLevel,
)
from docwen_plugin_markdown.to_docx.numbering import (
    _NUMBERING_XML_TEMPLATE,
    NUMBERING_CONTENT_TYPE,
    NUMBERING_REL_TYPE,
    WML_NS,
    _ensure_content_type,
    _ensure_relationship,
)

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════════
# Constants
# ═══════════════════════════════════════════════════════════════════════════

BASE_INDENT = 420  # twips
INDENT_INCREMENT = 420  # twips per level


def _bind_heading_style_ids(
    translation_result: TranslationResult,
    heading_style_ids: Mapping[int, str],
) -> TranslationResult:
    """Bind abstract heading levels to request-owned template style IDs."""

    levels: list[WordNumberingLevel] = []
    for level in translation_result.levels:
        heading_level = level.ilvl + 1
        style_id = heading_style_ids.get(heading_level)
        if not style_id:
            raise ValueError(f"Missing managed heading style binding for level {heading_level}.")
        levels.append(replace(level, p_style=style_id))
    return replace(translation_result, levels=levels)


# ═══════════════════════════════════════════════════════════════════════════
# OOXML element factories
# ═══════════════════════════════════════════════════════════════════════════


def create_heading_abstract_num(
    abstract_num_id: int,
    translation_result: TranslationResult,
) -> etree._Element:
    """Create a ``w:abstractNum`` element from the translation result's levels.

    Each level includes a ``w:pStyle`` linking the numbering to a heading
    style (e.g. Heading1), which is the key mechanism for Word-native
    heading numbering.

    Uses ``hybridMultilevel`` as *multiLevelType* (required for style-linked
    lists).

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

    # multiLevelType — must be hybridMultilevel for style-linked lists
    mlt = etree.SubElement(an, f"{{{WML_NS}}}multiLevelType")
    mlt.set(f"{{{WML_NS}}}val", "hybridMultilevel")

    for level in translation_result.levels:
        an.append(_create_heading_lvl(level))

    return an


def _create_heading_lvl(level: WordNumberingLevel) -> etree._Element:
    """Create a single ``w:lvl`` element from a ``WordNumberingLevel``."""
    lvl = etree.Element(f"{{{WML_NS}}}lvl")
    lvl.set(f"{{{WML_NS}}}ilvl", str(level.ilvl))

    # start
    e = etree.SubElement(lvl, f"{{{WML_NS}}}start")
    e.set(f"{{{WML_NS}}}val", level.start)

    # numFmt
    e = etree.SubElement(lvl, f"{{{WML_NS}}}numFmt")
    e.set(f"{{{WML_NS}}}val", level.num_fmt)

    # pStyle — THE KEY: links to heading style (e.g. Heading1)
    e = etree.SubElement(lvl, f"{{{WML_NS}}}pStyle")
    e.set(f"{{{WML_NS}}}val", level.p_style)

    # suff — what follows the number
    e = etree.SubElement(lvl, f"{{{WML_NS}}}suff")
    e.set(f"{{{WML_NS}}}val", level.suff)

    # lvlText — the format template (e.g. "%1、")
    e = etree.SubElement(lvl, f"{{{WML_NS}}}lvlText")
    e.set(f"{{{WML_NS}}}val", level.lvl_text)

    # lvlJc
    e = etree.SubElement(lvl, f"{{{WML_NS}}}lvlJc")
    e.set(f"{{{WML_NS}}}val", "left")

    # pPr → ind (hanging indent for clean wrapping)
    pPr = etree.SubElement(lvl, f"{{{WML_NS}}}pPr")
    ind = etree.SubElement(pPr, f"{{{WML_NS}}}ind")
    left_val = BASE_INDENT + INDENT_INCREMENT * level.ilvl
    ind.set(f"{{{WML_NS}}}left", str(left_val))
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
# Styles.xml injection
# ═══════════════════════════════════════════════════════════════════════════

# CT_PPrBase child order — numPr comes after these elements:
_PPR_BEFORE_NUMPR = {
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
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


def inject_numpr_into_heading_styles(
    styles_root: etree._Element,
    num_id: int,
    translation_result: TranslationResult,
) -> None:
    """Add ``w:numPr`` to each Heading style's ``w:pPr``.

    For each ``WordNumberingLevel`` in the translation result, finds the
    corresponding ``w:style`` element (e.g. Heading1) and injects or
    replaces its ``w:numPr``.

    This is the second binding path (in addition to ``pStyle`` in the lvl
    definition). Word writes both; the numPr on style takes priority per
    the OOXML spec.
    """
    for level in translation_result.levels:
        style_id = level.p_style  # e.g. "Heading1"
        style_elem = styles_root.find(f'{{{WML_NS}}}style[@{{{WML_NS}}}styleId="{style_id}"]')
        if style_elem is None:
            logger.warning("Style %s not found in styles.xml — skipping", style_id)
            continue

        # Find or create pPr — insert at position 0 per CT_Style ordering
        # (pPr must come before rPr, tblPr, etc.)
        pPr = style_elem.find(f"{{{WML_NS}}}pPr")
        if pPr is None:
            pPr = etree.Element(f"{{{WML_NS}}}pPr")
            style_elem.insert(0, pPr)

        # Remove existing numPr if present
        existing = pPr.find(f"{{{WML_NS}}}numPr")
        if existing is not None:
            pPr.remove(existing)

        # Build w:numPr
        numPr = etree.Element(f"{{{WML_NS}}}numPr")
        ilvl_elem = etree.SubElement(numPr, f"{{{WML_NS}}}ilvl")
        ilvl_elem.set(f"{{{WML_NS}}}val", str(level.ilvl))
        numId_elem = etree.SubElement(numPr, f"{{{WML_NS}}}numId")
        numId_elem.set(f"{{{WML_NS}}}val", str(num_id))

        # Insert numPr at the correct position within pPr
        _insert_numpr(pPr, numPr)


# ═══════════════════════════════════════════════════════════════════════════
# Helpers — numbering.xml element ordering
# ═══════════════════════════════════════════════════════════════════════════


def _allocate_ids(num_root: etree._Element) -> tuple[int, int]:
    """Find free abstractNumId and numId in an existing numbering root.

    Scans all existing ``abstractNumId`` and ``numId`` attribute values,
    then picks the next available integer starting at 20 000 (above the
    list-numbering range which starts at 10 000).

    Returns:
        ``(abstract_num_id, num_id)`` — typically the same number, since
        Word allows abstractNumId and numId to share values.
    """
    used: set[int] = set()
    for tag in ("abstractNumId", "numId"):
        for elem in num_root.iter(f"{{{WML_NS}}}{tag}"):
            val = elem.get(f"{{{WML_NS}}}val")
            if val is not None:
                with contextlib.suppress(ValueError):
                    used.add(int(val))
    candidate = 20000
    while candidate in used:
        candidate += 1
    return candidate, candidate


def _insert_abstract_num(
    num_root: etree._Element,
    abstract_num_elem: etree._Element,
) -> None:
    """Insert abstractNum before the first num element (OOXML schema order).

    CT_Numbering requires: all ``numPicBullet`` → all ``abstractNum`` → all
    ``num``.  If ``num`` elements already exist (from template or list
    numbering writeback), appending an ``abstractNum`` after them is
    schema-invalid.
    """
    first_num = num_root.find(f"{{{WML_NS}}}num")
    if first_num is not None:
        first_num.addprevious(abstract_num_elem)
    else:
        num_root.append(abstract_num_elem)


def _insert_num(
    num_root: etree._Element,
    num_elem: etree._Element,
) -> None:
    """Append num at the end (after all existing nums)."""
    num_root.append(num_elem)


# ═══════════════════════════════════════════════════════════════════════════
# Post-save ZIP writeback
# ═══════════════════════════════════════════════════════════════════════════


def write_heading_numbering_to_docx(
    docx_path: str,
    translation_result: TranslationResult,
    *,
    heading_style_ids: Mapping[int, str] | None = None,
) -> None:
    """Inject heading numbering definitions into a saved DOCX file.

    Opens *docx_path* as a ZIP, ensures ``word/numbering.xml`` exists,
    appends the heading ``abstractNum`` and ``num`` elements, injects
    ``numPr`` into Heading styles in ``word/styles.xml``, updates
    relationships and content types, then writes back.

    Must be called **after** ``Document.save()`` so the ZIP contains
    all body content.

    IDs are dynamically allocated starting at 20 000 to avoid conflicts
    with template numbering and the list-numbering range (10 000+).
    """
    if heading_style_ids is not None:
        translation_result = _bind_heading_style_ids(translation_result, heading_style_ids)

    if not translation_result.levels:
        logger.debug("No heading numbering levels to inject — skipping")
        return

    original = Path(docx_path)

    with tempfile.TemporaryDirectory(prefix=".dw-heading-numbering-", dir=original.parent) as tmpdir:
        tmp_path = Path(tmpdir) / "heading_numbering_writeback.docx"

        with (
            zipfile.ZipFile(str(original), "r") as zf_in,
            zipfile.ZipFile(str(tmp_path), "w", zipfile.ZIP_DEFLATED) as zf_out,
        ):
            names = set(zf_in.namelist())

            # ── Guard against missing critical files (BUG-13) ───────────
            if "word/_rels/document.xml.rels" not in names:
                logger.error("DOCX missing word/_rels/document.xml.rels — cannot inject heading numbering")
                return
            if "[Content_Types].xml" not in names:
                logger.error("DOCX missing [Content_Types].xml — cannot inject heading numbering")
                return

            # Read relationships and content types once
            rels_raw = zf_in.read("word/_rels/document.xml.rels")
            ct_raw = zf_in.read("[Content_Types].xml")
            rels_root = etree.fromstring(rels_raw)
            ct_root = etree.fromstring(ct_raw)

            has_numbering = "word/numbering.xml" in names
            has_styles = "word/styles.xml" in names

            # ── numbering.xml ───────────────────────────────────────────
            if has_numbering:
                raw = zf_in.read("word/numbering.xml")
                num_root = etree.fromstring(raw)
            else:
                num_root = etree.fromstring(_NUMBERING_XML_TEMPLATE)

            # Dynamically allocate IDs (BUG-2)
            abstract_num_id, num_id = _allocate_ids(num_root)

            # Build heading numbering elements with allocated IDs
            abstract_num_elem = create_heading_abstract_num(
                abstract_num_id,
                translation_result,
            )
            num_elem = create_heading_num(num_id, abstract_num_id)

            # Insert respecting OOXML schema order (BUG-3)
            _insert_abstract_num(num_root, abstract_num_elem)
            _insert_num(num_root, num_elem)

            num_bytes = etree.tostring(
                num_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

            if not has_numbering:
                _ensure_relationship(rels_root, "numbering.xml", NUMBERING_REL_TYPE)
                _ensure_content_type(ct_root, "/word/numbering.xml", NUMBERING_CONTENT_TYPE)

            # ── styles.xml: inject numPr into heading styles ────────────
            styles_bytes: bytes | None = None
            if has_styles:
                styles_raw = zf_in.read("word/styles.xml")
                styles_root = etree.fromstring(styles_raw)
                inject_numpr_into_heading_styles(
                    styles_root,
                    num_id,
                    translation_result,
                )
                styles_bytes = etree.tostring(
                    styles_root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

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
                elif item.filename == "word/styles.xml" and styles_bytes is not None:
                    zf_out.writestr(item, styles_bytes)
                else:
                    zf_out.writestr(item, zf_in.read(item.filename))

            if not has_numbering:
                # Write new numbering.xml with proper ZipInfo (BUG-14)
                zi = zipfile.ZipInfo("word/numbering.xml")
                zi.compress_type = zipfile.ZIP_DEFLATED
                zf_out.writestr(zi, num_bytes)

        tmp_path.replace(original)
        logger.info(
            "Heading numbering definitions injected: %d levels, numId=%s",
            len(translation_result.levels),
            num_id,
        )
