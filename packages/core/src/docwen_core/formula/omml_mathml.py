"""Native Office Math to MathML without silently flattening unknown operators."""

from __future__ import annotations

import logging

import lxml.etree as etree

from docwen_core.formula.constants import MATHML_NS, NAMED_FUNCTIONS, NARY_OP_MAP, OMML_NS, SYMBOL_MAP

logger = logging.getLogger(__name__)
M_NS = OMML_NS["m"]


def _child(node, name: str):
    return node.find(f"{{{M_NS}}}{name}")


def _property(node, group: str, name: str, default: str) -> str:
    properties = _child(node, group)
    value = _child(properties, name) if properties is not None else None
    present_default = "1" if name in {"nor", "degHide", "subHide", "supHide"} else default
    return value.get(f"{{{M_NS}}}val", present_default) if value is not None else default


def _contents(source, destination) -> None:
    for child in source:
        _convert_omml_to_mathml_node(child, destination)


def _argument(node, name: str, parent):
    source = _child(node, name)
    if source is None:
        raise ValueError(f"{etree.QName(node).localname} is missing {name}")
    row = etree.SubElement(parent, "mrow")
    _contents(source, row)
    return row


def _grouped_argument(node, parent) -> None:
    # Empty delimiters express TeX grouping, not visible parentheses. Keeping
    # the argument grouped prevents an imported function or n-ary expression
    # from changing its operand on the next export.
    group = etree.SubElement(parent, "mfenced", open="", close="", separators="")
    _argument(node, "e", group)


def _convert_omml_to_mathml_node(node, parent) -> None:
    name = etree.QName(node)
    tag = name.localname
    if name.namespace != M_NS:
        # Word run properties are presentation metadata inside a math run.
        if name.namespace == OMML_NS["w"] and tag in {"rPr", "rStyle"}:
            return
        raise ValueError(f"unsupported math namespace: {name.namespace}")
    if tag.endswith("Pr") or tag in {"ctrlPr", "argPr"}:
        return
    if tag in {"oMath", "oMathPara", "e", "fName"}:
        _contents(node, parent)
    elif tag == "r":
        leaf = _child(node, "t")
        if leaf is None or leaf.text is None:
            return
        text = leaf.text
        normal = _property(node, "rPr", "nor", "0") not in {"0", "false", "off"}
        if normal:
            native_tag = "mtext"
        elif text in SYMBOL_MAP or text in NARY_OP_MAP or not text.isalnum():
            native_tag = "mo"
        elif text.isdigit():
            native_tag = "mn"
        else:
            native_tag = "mi"
        output = etree.SubElement(parent, native_tag)
        output.text = text
        style = _property(node, "rPr", "sty", "")
        if not normal and style in {"i", "b", "bi"}:
            output.set("mathvariant", {"i": "italic", "b": "bold", "bi": "bold-italic"}[style])
        elif not normal and style == "p" and text.isalnum() and not text.isdigit() and text not in NAMED_FUNCTIONS:
            output.set("mathvariant", "normal")
    elif tag in {"f", "sSup", "sSub", "sSubSup", "limLow", "limUpp"}:
        native_tag, slots = {
            "f": ("mfrac", ("num", "den")),
            "sSup": ("msup", ("e", "sup")),
            "sSub": ("msub", ("e", "sub")),
            "sSubSup": ("msubsup", ("e", "sub", "sup")),
            "limLow": ("munder", ("e", "lim")),
            "limUpp": ("mover", ("e", "lim")),
        }[tag]
        output = etree.SubElement(parent, native_tag)
        if tag == "f":
            fraction_type = _property(node, "fPr", "type", "bar")
            if fraction_type == "noBar":
                output.set("linethickness", "0")
            elif fraction_type != "bar":
                raise ValueError(f"unsupported native fraction type: {fraction_type!r}")
        for slot in slots:
            _argument(node, slot, output)
    elif tag == "func":
        function = etree.SubElement(parent, "mrow")
        _argument(node, "fName", function)
        _grouped_argument(node, function)
    elif tag == "nary":
        operator = _property(node, "naryPr", "chr", "∫")
        limits_under = _property(node, "naryPr", "limLoc", "subSup") == "undOvr"
        lower = _child(node, "sub")
        upper = _child(node, "sup")
        show_lower = (
            lower is not None and len(lower) > 0 and _property(node, "naryPr", "subHide", "0") in {"0", "false", "off"}
        )
        show_upper = (
            upper is not None and len(upper) > 0 and _property(node, "naryPr", "supHide", "0") in {"0", "false", "off"}
        )
        if show_lower and show_upper:
            native_tag = "munderover" if limits_under else "msubsup"
        elif show_lower:
            native_tag = "munder" if limits_under else "msub"
        elif show_upper:
            native_tag = "mover" if limits_under else "msup"
        else:
            native_tag = "mrow"
        output = etree.SubElement(parent, native_tag)
        etree.SubElement(output, "mo").text = operator
        if show_lower:
            _argument(node, "sub", output)
        if show_upper:
            _argument(node, "sup", output)
        _grouped_argument(node, parent)
    elif tag == "rad":
        degree = _child(node, "deg")
        hidden = _property(node, "radPr", "degHide", "0") not in {"0", "false", "off"}
        indexed = not hidden and degree is not None and len(degree) > 0
        output = etree.SubElement(parent, "mroot" if indexed else "msqrt")
        _argument(node, "e", output)
        if indexed:
            _argument(node, "deg", output)
    elif tag == "m":
        output = etree.SubElement(parent, "mtable")
        for row in node:
            if etree.QName(row).localname == "mPr":
                continue
            if etree.QName(row).localname != "mr":
                raise ValueError("unsupported matrix row")
            native_row = etree.SubElement(output, "mtr")
            for cell in row:
                if etree.QName(cell).localname != "e":
                    raise ValueError("unsupported matrix cell")
                _contents(cell, etree.SubElement(native_row, "mtd"))
    elif tag == "d":
        output = etree.SubElement(
            parent,
            "mfenced",
            open=_property(node, "dPr", "begChr", "("),
            close=_property(node, "dPr", "endChr", ")"),
            separators=_property(node, "dPr", "sepChr", "|"),
        )
        for child in node:
            if etree.QName(child).localname == "e":
                _contents(child, etree.SubElement(output, "mrow"))
            elif etree.QName(child).localname != "dPr":
                raise ValueError("unsupported delimiter child")
    else:
        raise ValueError(f"unsupported OMML structure: {tag}")


def omml_to_mathml(omml_str: str) -> str | None:
    """Convert one complete supported formula, refusing lossy flattening."""
    try:
        source = etree.fromstring(omml_str.encode("utf-8"))
        math = etree.Element("math", xmlns=MATHML_NS)
        _convert_omml_to_mathml_node(source, math)
        if not len(math):
            raise ValueError("empty OMML expression")
        return etree.tostring(math, encoding="unicode")
    except Exception as exc:
        logger.error("OMML to MathML failed: %s", exc)
        return None
