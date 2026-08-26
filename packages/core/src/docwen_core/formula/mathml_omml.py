"""MathML to OMML conversion."""

from __future__ import annotations

import logging

import lxml.etree as etree

from docwen_core.formula.constants import BRACKET_PAIRS, OMML_NS

logger = logging.getLogger(__name__)

M_NS = OMML_NS["m"]


def _mk_omml(tag: str) -> str:
    """Create an OMML-qualified element tag string."""
    return f"{{{M_NS}}}{tag}"


def _convert_mathml_to_omml_node(mathml_node, omml_parent) -> None:
    """Recursively convert MathML node to OMML nodes."""
    tag = etree.QName(mathml_node).localname

    if tag == "math":
        children = list(mathml_node)
        _process_children_to_omml(children, omml_parent)

    elif tag in ("mi", "mn", "mo"):
        run = etree.SubElement(omml_parent, _mk_omml("r"))
        text_elem = etree.SubElement(run, _mk_omml("t"))
        text_elem.text = mathml_node.text or ""

    elif tag == "mfrac":
        frac = etree.SubElement(omml_parent, _mk_omml("f"))
        num = etree.SubElement(frac, _mk_omml("num"))
        den = etree.SubElement(frac, _mk_omml("den"))
        children = list(mathml_node)
        if len(children) >= 2:
            _convert_mathml_to_omml_node(children[0], num)
            _convert_mathml_to_omml_node(children[1], den)

    elif tag == "msup":
        ssup = etree.SubElement(omml_parent, _mk_omml("sSup"))
        base = etree.SubElement(ssup, _mk_omml("e"))
        sup = etree.SubElement(ssup, _mk_omml("sup"))
        children = list(mathml_node)

        if len(children) > 2:
            logger.debug(f"msup non-standard structure: {len(children)} children, using enhanced handling")
            _process_children_to_omml(children[:-1], base)
            _convert_mathml_to_omml_node(children[-1], sup)
        elif len(children) == 2:
            _convert_mathml_to_omml_node(children[0], base)
            _convert_mathml_to_omml_node(children[1], sup)

    elif tag == "msub":
        ssub = etree.SubElement(omml_parent, _mk_omml("sSub"))
        base = etree.SubElement(ssub, _mk_omml("e"))
        sub = etree.SubElement(ssub, _mk_omml("sub"))
        children = list(mathml_node)

        if len(children) > 2:
            logger.debug(f"msub non-standard structure: {len(children)} children, using enhanced handling")
            _process_children_to_omml(children[:-1], base)
            _convert_mathml_to_omml_node(children[-1], sub)
        elif len(children) == 2:
            _convert_mathml_to_omml_node(children[0], base)
            _convert_mathml_to_omml_node(children[1], sub)

    elif tag == "msubsup":
        subsup = etree.SubElement(omml_parent, _mk_omml("sSubSup"))
        base = etree.SubElement(subsup, _mk_omml("e"))
        sub = etree.SubElement(subsup, _mk_omml("sub"))
        sup = etree.SubElement(subsup, _mk_omml("sup"))
        children = list(mathml_node)

        if len(children) > 3:
            logger.debug(f"msubsup non-standard structure: {len(children)} children, using enhanced handling")
            _process_children_to_omml(children[:-2], base)
            _convert_mathml_to_omml_node(children[-2], sub)
            _convert_mathml_to_omml_node(children[-1], sup)
        elif len(children) >= 3:
            _convert_mathml_to_omml_node(children[0], base)
            _convert_mathml_to_omml_node(children[1], sub)
            _convert_mathml_to_omml_node(children[2], sup)

    elif tag == "msqrt":
        rad = etree.SubElement(omml_parent, _mk_omml("rad"))
        deg = etree.SubElement(rad, _mk_omml("deg"))
        base = etree.SubElement(rad, _mk_omml("e"))
        for child in mathml_node:
            _convert_mathml_to_omml_node(child, base)

    elif tag == "mroot":
        rad = etree.SubElement(omml_parent, _mk_omml("rad"))
        deg = etree.SubElement(rad, _mk_omml("deg"))
        base = etree.SubElement(rad, _mk_omml("e"))
        children = list(mathml_node)
        if len(children) >= 2:
            _convert_mathml_to_omml_node(children[0], base)
            _convert_mathml_to_omml_node(children[1], deg)

    elif tag == "mrow":
        for child in mathml_node:
            _convert_mathml_to_omml_node(child, omml_parent)

    elif tag == "mtable":
        m_elem = etree.SubElement(omml_parent, _mk_omml("m"))
        for mtr in mathml_node:
            if etree.QName(mtr).localname == "mtr":
                mr = etree.SubElement(m_elem, _mk_omml("mr"))
                for mtd in mtr:
                    if etree.QName(mtd).localname == "mtd":
                        e_elem = etree.SubElement(mr, _mk_omml("e"))
                        for child in mtd:
                            _convert_mathml_to_omml_node(child, e_elem)

    elif tag == "mfenced":
        open_char = mathml_node.get("open", "(")
        close_char = mathml_node.get("close", ")")

        d_elem = etree.SubElement(omml_parent, _mk_omml("d"))
        dPr = etree.SubElement(d_elem, _mk_omml("dPr"))
        begChr = etree.SubElement(dPr, _mk_omml("begChr"))
        begChr.set(_mk_omml("val"), open_char)
        endChr = etree.SubElement(dPr, _mk_omml("endChr"))
        endChr.set(_mk_omml("val"), close_char)

        e_elem = etree.SubElement(d_elem, _mk_omml("e"))
        for child in mathml_node:
            _convert_mathml_to_omml_node(child, e_elem)

    else:
        # Unknown tag: recurse into children
        for child in mathml_node:
            _convert_mathml_to_omml_node(child, omml_parent)


def _process_children_to_omml(children: list, omml_parent) -> None:
    """Process MathML child nodes, recognizing bracket-content-bracket patterns.

    Scans children for mo(left bracket) + content + mo(right bracket) patterns
    and converts them to OMML <m:d> (delimiter) structures.
    """
    i = 0
    while i < len(children):
        child = children[i]
        child_tag = etree.QName(child).localname

        if child_tag == "mo":
            open_char = (child.text or "").strip()

            if open_char in BRACKET_PAIRS:
                expected_close = BRACKET_PAIRS[open_char]

                # Look for matching closing bracket
                close_idx = -1
                for j in range(i + 2, len(children)):
                    check_child = children[j]
                    if etree.QName(check_child).localname == "mo":
                        close_text = (check_child.text or "").strip()
                        if close_text == expected_close:
                            close_idx = j
                            break

                if close_idx > i + 1:
                    logger.debug(
                        f"Detected bracket pattern: {open_char}...{expected_close}, content nodes: {close_idx - i - 1}"
                    )

                    d_elem = etree.SubElement(omml_parent, _mk_omml("d"))
                    dPr = etree.SubElement(d_elem, _mk_omml("dPr"))
                    begChr = etree.SubElement(dPr, _mk_omml("begChr"))
                    begChr.set(_mk_omml("val"), open_char)
                    endChr = etree.SubElement(dPr, _mk_omml("endChr"))
                    endChr.set(_mk_omml("val"), expected_close)

                    e_elem = etree.SubElement(d_elem, _mk_omml("e"))
                    for content_idx in range(i + 1, close_idx):
                        _convert_mathml_to_omml_node(children[content_idx], e_elem)

                    i = close_idx + 1
                    continue

        _convert_mathml_to_omml_node(child, omml_parent)
        i += 1


def mathml_to_omml(mathml_str: str) -> str | None:
    """Convert MathML to OMML format.

    Args:
        mathml_str: MathML XML string.

    Returns:
        OMML XML string, or None if conversion fails.
    """
    try:
        mathml_tree = etree.fromstring(mathml_str.encode("utf-8"))
        omml_math = etree.Element(_mk_omml("oMath"), nsmap={"m": OMML_NS["m"]})
        _convert_mathml_to_omml_node(mathml_tree, omml_math)
        return etree.tostring(omml_math, encoding="unicode", pretty_print=True)
    except Exception as e:
        logger.error(f"MathML to OMML failed: {e}")
        return None
