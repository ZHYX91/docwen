"""MathML to native Office Math, with explicit operator arguments."""

from __future__ import annotations

import logging
from typing import Any

import lxml.etree as etree

from docwen_core.formula.constants import BRACKET_PAIRS, LIMIT_FUNCTIONS, NAMED_FUNCTIONS, NARY_OP_MAP, OMML_NS
from docwen_core.formula.fractions import mathml_fraction_type

logger = logging.getLogger(__name__)
M_NS = OMML_NS["m"]
_ARGUMENT_SEPARATORS = {"+", "-", "−", "=", "<", ">", "≤", "≥", "≠", "≈", "→", ",", ";"}


def _mk_omml(tag: str) -> str:
    return f"{{{M_NS}}}{tag}"


def _tag(node) -> str:
    return etree.QName(node).localname


def _property(parent, tag: str, value: str) -> None:
    etree.SubElement(parent, _mk_omml(tag)).set(_mk_omml("val"), value)


def _children(node, expected: int | None = None) -> list:
    children = list(node)
    if expected is not None and len(children) != expected:
        raise ValueError(f"{_tag(node)} requires {expected} children, found {len(children)}")
    return children


def _atom(node):
    while _tag(node) == "mrow" and len(node) == 1:
        node = node[0]
    return node


def _operator_parts(node) -> tuple[Any, Any | None, Any | None] | None:
    """Read an explicit operator token and its scripts, never ordinary text."""
    node = _atom(node)
    tag = _tag(node)
    lower = upper = None
    if tag in {"msub", "munder", "msup", "mover", "msubsup", "munderover"}:
        children = _children(node, 3 if tag in {"msubsup", "munderover"} else 2)
        node = _atom(children[0])
        if tag in {"msub", "munder", "msubsup", "munderover"}:
            lower = children[1]
        if tag in {"msup", "mover"}:
            upper = children[1]
        elif tag in {"msubsup", "munderover"}:
            upper = children[2]
    token = node.text or ""
    if _tag(node) == "mo" and (token in NARY_OP_MAP or token in NAMED_FUNCTIONS):
        return node, lower, upper
    if _tag(node) == "mi" and token in NAMED_FUNCTIONS and node.get("mathvariant") not in {"italic", "bold-italic"}:
        return node, lower, upper
    return None


def _delimiter_end(children: list, start: int) -> int | None:
    """Balance nested delimiter tokens instead of taking the first closer."""
    opening = children[start]
    if _tag(opening) != "mo" or (opening.text or "") not in BRACKET_PAIRS:
        return None
    stack = [BRACKET_PAIRS[opening.text]]
    for index in range(start + 1, len(children)):
        child = children[index]
        if _tag(child) != "mo":
            continue
        text = child.text or ""
        if text == stack[-1]:
            stack.pop()
            if not stack:
                return index
        elif text in BRACKET_PAIRS:
            stack.append(BRACKET_PAIRS[text])
    return None


def _product_end(children: list, start: int) -> int:
    """Keep an n-ary operand's factors together up to a sum/relation.

    Presentation MathML has no apply node. Sibling factors are one argument;
    grouped expressions, fractions and nested delimiters remain indivisible.
    Explicit grouping is required to apply an operator to a whole sum.
    """
    index = start
    while index < len(children):
        child = children[index]
        if index > start and _tag(child) == "mo" and (child.text or "") in _ARGUMENT_SEPARATORS:
            break
        closing = _delimiter_end(children, index)
        index = closing + 1 if closing is not None else index + 1
    return index


def _limit_argument_end(children: list, start: int) -> int:
    """Require one structurally bounded operand for a native limit function.

    Presentation MathML cannot establish the scope of ungrouped function
    tokens. Guessing would turn products into function arguments or move
    factors outside a limit. Keep the complete authored formula on failure.
    """
    if start >= len(children):
        raise ValueError("limit operator has no argument")
    closing = _delimiter_end(children, start)
    if closing is not None:
        end = closing + 1
    elif _tag(children[start]) in {"mrow", "mfrac", "mfenced"}:
        end = start + 1
    else:
        raise ValueError("ambiguous limit operand requires an explicit group or fraction")
    if end < len(children):
        following = children[end]
        if _tag(following) != "mo" or (following.text or "") not in _ARGUMENT_SEPARATORS:
            raise ValueError("ambiguous factors after a limit operand require explicit grouping")
    return end


def _append_operator(parts, arguments: list, parent) -> None:
    operator, lower, upper = parts
    token = operator.text or ""
    if not arguments:
        raise ValueError(f"operator {token!r} has no supported argument")
    if token in NARY_OP_MAP:
        native = etree.SubElement(parent, _mk_omml("nary"))
        properties = etree.SubElement(native, _mk_omml("naryPr"))
        _property(properties, "chr", token)
        _property(properties, "limLoc", "undOvr" if token in {"∑", "∏", "⋃", "⋂"} else "subSup")
        _property(properties, "subHide", "1" if lower is None else "0")
        _property(properties, "supHide", "1" if upper is None else "0")
        for name, script in (("sub", lower), ("sup", upper)):
            destination = etree.SubElement(native, _mk_omml(name))
            if script is not None:
                _convert_mathml_to_omml_node(script, destination)
        argument = etree.SubElement(native, _mk_omml("e"))
    else:
        native = etree.SubElement(parent, _mk_omml("func"))
        name = etree.SubElement(native, _mk_omml("fName"))
        target = name
        # A limit plus a following fraction must share m:func. A bare limLow
        # preceding m:f is regrouped into its numerator by LibreOffice.
        if upper is not None:
            limit = etree.SubElement(target, _mk_omml("limUpp"))
            target = etree.SubElement(limit, _mk_omml("e"))
            _convert_mathml_to_omml_node(upper, etree.SubElement(limit, _mk_omml("lim")))
        if lower is not None:
            limit = etree.SubElement(target, _mk_omml("limLow"))
            target = etree.SubElement(limit, _mk_omml("e"))
            _convert_mathml_to_omml_node(lower, etree.SubElement(limit, _mk_omml("lim")))
        _convert_mathml_to_omml_node(operator, target)
        argument = etree.SubElement(native, _mk_omml("e"))
    _process_children_to_omml(arguments, argument)


def _convert_mathml_to_omml_node(node, parent) -> None:
    tag = _tag(node)
    if tag in {"math", "mrow"}:
        _process_children_to_omml(_children(node), parent)
    elif tag in {"mi", "mn", "mo", "mtext"}:
        text = node.text or ""
        if tag != "mtext" and text.startswith("\\") and len(text) > 1:
            raise ValueError(f"unsupported LaTeX command: {text}")
        run = etree.SubElement(parent, _mk_omml("r"))
        if tag == "mtext":
            properties = etree.SubElement(run, _mk_omml("rPr"))
            etree.SubElement(properties, _mk_omml("nor"))
        else:
            variant = node.get("mathvariant")
            style = {"normal": "p", "bold": "b", "italic": "i", "bold-italic": "bi"}.get(variant)
            if style is None and (tag == "mo" or text in NAMED_FUNCTIONS):
                style = "p"
            if variant is not None and style is None:
                raise ValueError(f"unsupported mathvariant: {variant}")
            if style is not None:
                _property(etree.SubElement(run, _mk_omml("rPr")), "sty", style)
        leaf = etree.SubElement(run, _mk_omml("t"))
        leaf.text = text
        if tag == "mtext" or any(character.isspace() for character in text):
            leaf.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    elif tag in {"mfrac", "msup", "msub", "msubsup"}:
        native_tag, slots = {
            "mfrac": ("f", ("num", "den")),
            "msup": ("sSup", ("e", "sup")),
            "msub": ("sSub", ("e", "sub")),
            "msubsup": ("sSubSup", ("e", "sub", "sup")),
        }[tag]
        children = _children(node, 3 if tag == "msubsup" else 2)
        native = etree.SubElement(parent, _mk_omml(native_tag))
        if tag == "mfrac" and mathml_fraction_type(node) == "noBar":
            _property(etree.SubElement(native, _mk_omml("fPr")), "type", "noBar")
        for child, slot in zip(children, slots, strict=True):
            _convert_mathml_to_omml_node(child, etree.SubElement(native, _mk_omml(slot)))
    elif tag in {"msqrt", "mroot"}:
        children = _children(node, 2 if tag == "mroot" else None)
        if not children:
            raise ValueError("radical has no radicand")
        radical = etree.SubElement(parent, _mk_omml("rad"))
        properties = etree.SubElement(radical, _mk_omml("radPr"))
        _property(properties, "degHide", "1" if tag == "msqrt" else "0")
        degree = etree.SubElement(radical, _mk_omml("deg"))
        radicand = etree.SubElement(radical, _mk_omml("e"))
        if tag == "mroot":
            _convert_mathml_to_omml_node(children[1], degree)
            _convert_mathml_to_omml_node(children[0], radicand)
        else:
            _process_children_to_omml(children, radicand)
    elif tag == "mtable":
        matrix = etree.SubElement(parent, _mk_omml("m"))
        for row in node:
            if _tag(row) != "mtr":
                raise ValueError(f"unsupported matrix row: {_tag(row)}")
            native_row = etree.SubElement(matrix, _mk_omml("mr"))
            for cell in row:
                if _tag(cell) != "mtd":
                    raise ValueError(f"unsupported matrix cell: {_tag(cell)}")
                _process_children_to_omml(list(cell), etree.SubElement(native_row, _mk_omml("e")))
    elif tag == "mfenced":
        native = etree.SubElement(parent, _mk_omml("d"))
        properties = etree.SubElement(native, _mk_omml("dPr"))
        _property(properties, "begChr", node.get("open", "("))
        _property(properties, "endChr", node.get("close", ")"))
        children = list(node)
        # Each argument owns a separator in MathML mfenced. One container
        # would silently drop those separators for multi-argument functions.
        if len(children) > 1:
            separators = node.get("separators", ",").replace(" ", "")
            if len(set(separators)) > 1:
                raise ValueError("mfenced with differing separators is unsupported")
            _property(properties, "sepChr", separators or "")
        for child in children:
            _convert_mathml_to_omml_node(child, etree.SubElement(native, _mk_omml("e")))
    elif tag == "mstyle" and not node.attrib:
        _process_children_to_omml(list(node), parent)
    else:
        # Flattening an unknown container can erase an accent, enclosure or
        # other mathematical operator while leaving a plausible expression.
        raise ValueError(f"unsupported MathML structure: {tag}")


def _process_children_to_omml(children: list, parent) -> None:
    index = 0
    while index < len(children):
        child = children[index]
        operator = _operator_parts(child)
        token = operator[0].text if operator is not None else None
        if operator is not None and (token in NARY_OP_MAP or token in LIMIT_FUNCTIONS):
            end = (
                _product_end(children, index + 1) if token in NARY_OP_MAP else _limit_argument_end(children, index + 1)
            )
            _append_operator(operator, children[index + 1 : end], parent)
            index = end
            continue
        # Ordinary named functions keep their source tokens and side scripts.
        # m:func would invent operand ownership absent from Presentation MathML.
        closing = _delimiter_end(children, index)
        if closing is not None:
            native = etree.SubElement(parent, _mk_omml("d"))
            properties = etree.SubElement(native, _mk_omml("dPr"))
            _property(properties, "begChr", child.text or "")
            _property(properties, "endChr", children[closing].text or "")
            _process_children_to_omml(children[index + 1 : closing], etree.SubElement(native, _mk_omml("e")))
            index = closing + 1
            continue
        _convert_mathml_to_omml_node(child, parent)
        index += 1


def mathml_to_omml(mathml_str: str) -> str | None:
    """Convert a supported complete expression; refuse partial conversion."""
    try:
        tree = etree.fromstring(mathml_str.encode("utf-8"))
        # Inspect every node before operator/delimiter optimizations skip it.
        # Unsupported inherited/token styles must not silently look successful.
        for node in tree.iter():
            variant = node.get("mathvariant")
            if variant is None:
                continue
            tag = _tag(node)
            if tag not in {"mi", "mn", "mo"} or variant not in {"normal", "bold", "italic", "bold-italic"}:
                if tag == "mtext" and variant == "normal":
                    continue
                raise ValueError(f"unsupported mathvariant on {tag}: {variant}")
        native = etree.Element(_mk_omml("oMath"), nsmap={"m": M_NS})
        _convert_mathml_to_omml_node(tree, native)
        if not len(native):
            raise ValueError("empty MathML expression")
        return etree.tostring(native, encoding="unicode")
    except Exception as exc:
        logger.error("MathML to OMML failed: %s", exc)
        return None
