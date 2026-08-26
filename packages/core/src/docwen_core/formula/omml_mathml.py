"""OMML to MathML conversion."""

from __future__ import annotations

import logging

import lxml.etree as etree

from docwen_core.formula.constants import MATHML_NS, NARY_OP_MAP, OMML_NS, SYMBOL_MAP

logger = logging.getLogger(__name__)

M_NS = OMML_NS["m"]


def _get_child_by_tag(element, tag_name: str):
    """Find direct child element by local tag name."""
    for child in element:
        if etree.QName(child).localname == tag_name:
            return child
    return None


def _convert_omml_to_mathml_node(omml_node, mathml_parent) -> None:
    """Recursively convert OMML node to MathML nodes."""
    tag = etree.QName(omml_node).localname

    if tag == "oMath":
        for child in omml_node:
            _convert_omml_to_mathml_node(child, mathml_parent)

    elif tag == "r":
        t_elem = _get_child_by_tag(omml_node, "t")

        if t_elem is not None and t_elem.text:
            text = t_elem.text.strip()
            # Determine MathML element type based on content
            if text in SYMBOL_MAP or text in NARY_OP_MAP or not text.isalnum():
                elem = etree.SubElement(mathml_parent, "mo")
            elif text.isdigit():
                elem = etree.SubElement(mathml_parent, "mn")
            else:
                elem = etree.SubElement(mathml_parent, "mi")
            elem.text = text

    elif tag == "f":
        frac = etree.SubElement(mathml_parent, "mfrac")
        num_elem = _get_child_by_tag(omml_node, "num")
        den_elem = _get_child_by_tag(omml_node, "den")

        if num_elem is not None:
            num_row = etree.SubElement(frac, "mrow")
            for child in num_elem:
                _convert_omml_to_mathml_node(child, num_row)
        else:
            etree.SubElement(frac, "mrow")

        if den_elem is not None:
            den_row = etree.SubElement(frac, "mrow")
            for child in den_elem:
                _convert_omml_to_mathml_node(child, den_row)
        else:
            etree.SubElement(frac, "mrow")

    elif tag == "sSup":
        ssup = etree.SubElement(mathml_parent, "msup")
        base_elem = _get_child_by_tag(omml_node, "e")
        sup_elem = _get_child_by_tag(omml_node, "sup")

        if base_elem is not None:
            base_row = etree.SubElement(ssup, "mrow")
            for child in base_elem:
                _convert_omml_to_mathml_node(child, base_row)
        else:
            etree.SubElement(ssup, "mrow")

        if sup_elem is not None:
            sup_row = etree.SubElement(ssup, "mrow")
            for child in sup_elem:
                _convert_omml_to_mathml_node(child, sup_row)
        else:
            etree.SubElement(ssup, "mrow")

    elif tag == "sSub":
        ssub = etree.SubElement(mathml_parent, "msub")
        base_elem = _get_child_by_tag(omml_node, "e")
        sub_elem = _get_child_by_tag(omml_node, "sub")

        if base_elem is not None:
            base_row = etree.SubElement(ssub, "mrow")
            for child in base_elem:
                _convert_omml_to_mathml_node(child, base_row)
        else:
            etree.SubElement(ssub, "mrow")

        if sub_elem is not None:
            sub_row = etree.SubElement(ssub, "mrow")
            for child in sub_elem:
                _convert_omml_to_mathml_node(child, sub_row)
        else:
            etree.SubElement(ssub, "mrow")

    elif tag == "sSubSup":
        subsup = etree.SubElement(mathml_parent, "msubsup")
        base_elem = _get_child_by_tag(omml_node, "e")
        sub_elem = _get_child_by_tag(omml_node, "sub")
        sup_elem = _get_child_by_tag(omml_node, "sup")

        if base_elem is not None:
            base_row = etree.SubElement(subsup, "mrow")
            for child in base_elem:
                _convert_omml_to_mathml_node(child, base_row)
        else:
            etree.SubElement(subsup, "mrow")

        if sub_elem is not None:
            sub_row = etree.SubElement(subsup, "mrow")
            for child in sub_elem:
                _convert_omml_to_mathml_node(child, sub_row)
        else:
            etree.SubElement(subsup, "mrow")

        if sup_elem is not None:
            sup_row = etree.SubElement(subsup, "mrow")
            for child in sup_elem:
                _convert_omml_to_mathml_node(child, sup_row)
        else:
            etree.SubElement(subsup, "mrow")

    elif tag == "nary":
        naryPr = _get_child_by_tag(omml_node, "naryPr")
        op_char = "∫"
        lim_loc = "subSup"

        if naryPr is not None:
            chr_elem = _get_child_by_tag(naryPr, "chr")
            if chr_elem is not None:
                op_char = chr_elem.attrib.get(f"{{{M_NS}}}val", "∫")

            lim_loc_elem = _get_child_by_tag(naryPr, "limLoc")
            if lim_loc_elem is not None:
                lim_loc = lim_loc_elem.attrib.get(f"{{{M_NS}}}val", "subSup")

        mathml_tag = "munderover" if lim_loc == "undOvr" else "msubsup"
        nary_elem = etree.SubElement(mathml_parent, mathml_tag)

        op = etree.SubElement(nary_elem, "mo")
        op.text = op_char

        sub_elem = _get_child_by_tag(omml_node, "sub")
        sub_row = etree.SubElement(nary_elem, "mrow")
        if sub_elem is not None:
            for child in sub_elem:
                _convert_omml_to_mathml_node(child, sub_row)

        sup_elem = _get_child_by_tag(omml_node, "sup")
        sup_row = etree.SubElement(nary_elem, "mrow")
        if sup_elem is not None:
            for child in sup_elem:
                _convert_omml_to_mathml_node(child, sup_row)

        base_elem = _get_child_by_tag(omml_node, "e")
        if base_elem is not None:
            for child in base_elem:
                _convert_omml_to_mathml_node(child, mathml_parent)

    elif tag == "rad":
        deg_elem = _get_child_by_tag(omml_node, "deg")
        base_elem = _get_child_by_tag(omml_node, "e")
        if deg_elem is not None:
            text = "".join(t for t in deg_elem.itertext()).strip()
            if text:
                root = etree.SubElement(mathml_parent, "mroot")
                if base_elem is not None:
                    for child in base_elem:
                        _convert_omml_to_mathml_node(child, root)
                for child in deg_elem:
                    _convert_omml_to_mathml_node(child, root)
                return

        sqrt = etree.SubElement(mathml_parent, "msqrt")
        if base_elem is not None:
            for child in base_elem:
                _convert_omml_to_mathml_node(child, sqrt)

    elif tag == "limLow":
        munder = etree.SubElement(mathml_parent, "munder")

        base_elem = _get_child_by_tag(omml_node, "e")
        lim_elem = _get_child_by_tag(omml_node, "lim")

        if base_elem is not None:
            base_row = etree.SubElement(munder, "mrow")
            for child in base_elem:
                _convert_omml_to_mathml_node(child, base_row)
        else:
            mo = etree.SubElement(munder, "mo")
            mo.text = "lim"

        if lim_elem is not None:
            lim_row = etree.SubElement(munder, "mrow")
            for child in lim_elem:
                _convert_omml_to_mathml_node(child, lim_row)
        else:
            etree.SubElement(munder, "mrow")

    elif tag == "m":
        mtable = etree.SubElement(mathml_parent, "mtable")

        for mr in omml_node:
            if etree.QName(mr).localname == "mr":
                mtr = etree.SubElement(mtable, "mtr")

                for e_elem in mr:
                    if etree.QName(e_elem).localname == "e":
                        mtd = etree.SubElement(mtr, "mtd")

                        for child in e_elem:
                            _convert_omml_to_mathml_node(child, mtd)

    elif tag == "d":
        dPr = _get_child_by_tag(omml_node, "dPr")
        open_char = "("
        close_char = ")"

        if dPr is not None:
            begChr = _get_child_by_tag(dPr, "begChr")
            if begChr is not None:
                open_char = begChr.get(f"{{{M_NS}}}val", "(")
            endChr = _get_child_by_tag(dPr, "endChr")
            if endChr is not None:
                close_char = endChr.get(f"{{{M_NS}}}val", ")")

        mfenced = etree.SubElement(mathml_parent, "mfenced")
        mfenced.set("open", open_char)
        mfenced.set("close", close_char)

        e_elem = _get_child_by_tag(omml_node, "e")
        if e_elem is not None:
            for child in e_elem:
                _convert_omml_to_mathml_node(child, mfenced)

    elif tag == "e":
        # OMML element container: pass through to parent
        for child in omml_node:
            _convert_omml_to_mathml_node(child, mathml_parent)

    else:
        # Unknown tag: recurse into children
        for child in omml_node:
            _convert_omml_to_mathml_node(child, mathml_parent)


def omml_to_mathml(omml_str: str) -> str | None:
    """Convert OMML to MathML format.

    Args:
        omml_str: OMML XML string.

    Returns:
        MathML XML string, or None if conversion fails.
    """
    try:
        omml_tree = etree.fromstring(omml_str.encode("utf-8"))
        mathml_math = etree.Element("math", xmlns=MATHML_NS)
        _convert_omml_to_mathml_node(omml_tree, mathml_math)
        return etree.tostring(mathml_math, encoding="unicode", pretty_print=True)
    except Exception as e:
        logger.error(f"OMML to MathML failed: {e}")
        return None
