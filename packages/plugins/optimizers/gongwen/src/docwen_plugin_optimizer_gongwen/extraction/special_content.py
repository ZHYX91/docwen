"""Detect special content in DOCX paragraphs (textboxes, tables, images, formulas, breaks)."""

from __future__ import annotations

from typing import TYPE_CHECKING

import lxml.etree as etree

from docwen_core.docx_parsing.break_utils import (
    detect_page_break,
    detect_section_break,
)
from docwen_core.docx_parsing.xml_ns import NS_M, NS_W, NS_WP

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph


def has_omml_formula(para: Paragraph) -> bool:
    """Check if paragraph contains OMML formula elements."""
    nsmap = {"m": NS_M, "w": NS_W}
    omaths = para._p.findall(".//m:oMath", nsmap)
    omath_paras = para._p.findall(".//m:oMathPara", nsmap)
    return len(omaths) > 0 or len(omath_paras) > 0


def has_image(para: Paragraph) -> bool:
    """Check if paragraph contains inline images/drawings."""
    nsmap = {"wp": NS_WP, "w": NS_W}
    drawings = para._p.findall(".//w:drawing", nsmap)
    return len(drawings) > 0


def has_page_break(para: Paragraph) -> bool:
    """Check if paragraph contains a page break."""
    return detect_page_break(para)


def has_section_break(para: Paragraph) -> bool:
    """Check if paragraph contains a section break."""
    return detect_section_break(para) is not None


def is_in_textbox(para: Paragraph, doc) -> bool:
    """Check if paragraph is inside a textbox.

    Traverses parent elements looking for textbox containers.
    """
    parent = para._p.getparent()
    while parent is not None:
        tag = etree.QName(parent).localname
        if tag in ("txbxContent", "textbox"):
            return True
        parent = parent.getparent()
    return False


def detect_table_context(para: Paragraph) -> str:
    """Detect if paragraph is inside a table cell.

    Returns:
        "header" if in a table header cell (w:tblHeader)
        "body" if in a regular table cell
        "" if not in a table
    """
    parent = para._p.getparent()

    # Check if inside a table cell
    tc = parent
    while tc is not None:
        tag = etree.QName(tc).localname
        if tag == "tc":
            # Found a table cell — check if it's a header
            tr = tc.getparent()
            if tr is not None:
                tr_tag = etree.QName(tr).localname
                if tr_tag == "tr":
                    # Check for tblHeader
                    tbl_pr = tr.find(f"{{{NS_W}}}trPr")
                    if tbl_pr is not None:
                        tbl_header = tbl_pr.find(f"{{{NS_W}}}tblHeader")
                        if tbl_header is not None:
                            return "header"
            return "body"
        tc = tc.getparent()

    return ""
