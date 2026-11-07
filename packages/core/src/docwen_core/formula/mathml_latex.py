"""MathML to LaTeX conversion."""

from __future__ import annotations

import logging
import re

import lxml.etree as etree

from docwen_core.formula.constants import LIMIT_FUNCTIONS, NARY_OP_MAP, SYMBOL_MAP
from docwen_core.formula.fractions import mathml_fraction_type

logger = logging.getLogger(__name__)


def _single_matrix(elem):
    """Recognize only a matrix that owns the complete group.

    Searching for any descendant table discarded surrounding operators and
    factors. Explicit outer delimiters must survive around a compound group.
    """
    while etree.QName(elem).localname == "mrow" and len(elem) == 1:
        elem = elem[0]
    return elem if etree.QName(elem).localname == "mtable" else None


def _convert_matrix_rows(mtable_elem) -> str:
    """Convert mtable element to LaTeX matrix rows."""
    rows: list[str] = []
    for mtr_elem in mtable_elem:
        if etree.QName(mtr_elem).localname != "mtr":
            raise ValueError("unsupported matrix row")
        cells: list[str] = []
        for mtd_elem in mtr_elem:
            if etree.QName(mtd_elem).localname != "mtd":
                raise ValueError("unsupported matrix cell")
            cell_content = _join_latex([_mathml_node_to_latex(c) for c in mtd_elem]).strip()
            cells.append(cell_content)
        if not cells:
            raise ValueError("matrix row has no cells")
        rows.append(" & ".join(cells))
    if not rows:
        raise ValueError("matrix has no rows")
    return " \\\\\n".join(rows)


def _join_latex(parts: list[str]) -> str:
    """Join generated atoms without extending a preceding TeX control word."""
    output: list[str] = []
    follows_control_word = False
    for part in parts:
        if not part:
            continue
        if follows_control_word and re.match(r"[A-Za-z]", part):
            output.append("{}")
        output.append(part)
        follows_control_word = re.search(r"(?<!\\)(?:\\\\)*\\[A-Za-z]+$", part) is not None
    return "".join(output)


def _text_to_latex(text: str) -> str:
    escapes = {
        "\\": r"\textbackslash{}",
        "{": r"\{",
        "}": r"\}",
        "$": r"\$",
        "&": r"\&",
        "%": r"\%",
        "#": r"\#",
        "_": r"\_",
        "^": r"\textasciicircum{}",
        "~": r"\textasciitilde{}",
    }
    return "\\text{" + "".join(escapes.get(character, character) for character in text) + "}"


def _math_literal(text: str) -> str:
    """Keep visible special characters distinct from TeX control syntax."""
    escapes = {"\\": r"\backslash", "{": r"\{", "}": r"\}", "$": r"\$", "&": r"\&", "%": r"\%", "#": r"\#", "_": r"\_"}
    if "^" in text or "~" in text:
        raise ValueError("unsupported literal math accent character")
    return _join_latex([escapes.get(character, character) for character in text])


def _children(node, count: int) -> list:
    children = list(node)
    if len(children) != count:
        raise ValueError(f"{etree.QName(node).localname} requires {count} children, found {len(children)}")
    return children


def _mathml_node_to_latex(node) -> str:
    """Recursively convert MathML node to LaTeX."""
    tag = etree.QName(node).localname

    if tag == "math":
        return _join_latex([_mathml_node_to_latex(child) for child in node])

    elif tag == "mtext":
        return _text_to_latex(node.text or "")

    elif tag in ("mi", "mn", "mo"):
        text = node.text or ""
        variant = node.get("mathvariant")
        if variant in {"normal", "italic", "bold", "bold-italic"}:
            command = {"normal": "mathrm", "italic": "mathit", "bold": "mathbf", "bold-italic": "boldsymbol"}[variant]
            return f"\\{command}{{{_math_literal(text)}}}"
        if variant is not None:
            raise ValueError(f"unsupported mathvariant: {variant}")
        return SYMBOL_MAP[text] if text in SYMBOL_MAP else _math_literal(text)

    elif tag == "mfrac":
        children = _children(node, 2)
        fraction_type = mathml_fraction_type(node)
        num = _mathml_node_to_latex(children[0]).strip()
        den = _mathml_node_to_latex(children[1]).strip()
        if fraction_type == "noBar":
            # The pinned LaTeX parser supports grouped \\atop, including
            # nested expressions; it does not support \\genfrac.
            return f"{{{num} \\atop {den}}}"
        return f"\\frac{{{num}}}{{{den}}}"

    elif tag == "msup":
        children = _children(node, 2)
        base = _mathml_node_to_latex(children[0]).strip()
        sup = _mathml_node_to_latex(children[1]).strip()
        if len(base) == 1 or base in NARY_OP_MAP.values():
            return f"{base}^{{{sup}}}"
        return f"{{{base}}}^{{{sup}}}"

    elif tag == "msub":
        children = _children(node, 2)
        base = _mathml_node_to_latex(children[0]).strip()
        sub = _mathml_node_to_latex(children[1]).strip()
        if len(base) == 1 or base in NARY_OP_MAP.values():
            return f"{base}_{{{sub}}}"
        return f"{{{base}}}_{{{sub}}}"

    elif tag == "msubsup":
        children = _children(node, 3)
        base = _mathml_node_to_latex(children[0]).strip()
        sub = _mathml_node_to_latex(children[1]).strip()
        sup = _mathml_node_to_latex(children[2]).strip()
        if base in NARY_OP_MAP.values() or len(base) == 1:
            return f"{base}_{{{sub}}}^{{{sup}}}"
        return f"{{{base}}}_{{{sub}}}^{{{sup}}}"

    elif tag == "munderover":
        children = _children(node, 3)
        base = _mathml_node_to_latex(children[0]).strip()
        under = _mathml_node_to_latex(children[1]).strip()
        over = _mathml_node_to_latex(children[2]).strip()
        if base in NARY_OP_MAP.values() or base in {SYMBOL_MAP[name] for name in LIMIT_FUNCTIONS}:
            return f"{base}_{{{under}}}^{{{over}}}"
        return f"\\overset{{{over}}}{{\\underset{{{under}}}{{{base}}}}}"

    elif tag == "munder":
        children = _children(node, 2)
        base = _mathml_node_to_latex(children[0]).strip()
        under = _mathml_node_to_latex(children[1]).strip()
        if base in NARY_OP_MAP.values() or base in {SYMBOL_MAP[name] for name in LIMIT_FUNCTIONS}:
            return f"{base}_{{{under}}}"
        return f"\\underset{{{under}}}{{{base}}}"

    elif tag == "mover":
        children = _children(node, 2)
        base = _mathml_node_to_latex(children[0]).strip()
        over = _mathml_node_to_latex(children[1]).strip()
        if base in NARY_OP_MAP.values() or base in {SYMBOL_MAP[name] for name in LIMIT_FUNCTIONS}:
            return f"{base}^{{{over}}}"
        return f"\\overset{{{over}}}{{{base}}}"

    elif tag == "msqrt":
        content = _join_latex([_mathml_node_to_latex(child) for child in node]).strip()
        return f"\\sqrt{{{content}}}"

    elif tag == "mroot":
        children = _children(node, 2)
        base = _mathml_node_to_latex(children[0]).strip()
        deg = _mathml_node_to_latex(children[1]).strip()
        return f"\\sqrt[{deg}]{{{base}}}"

    elif tag == "mrow":
        children = list(node)
        parts: list[str] = []
        i = 0

        while i < len(children):
            child = children[i]
            child_tag = etree.QName(child).localname
            part = _mathml_node_to_latex(child)

            if part:
                need_space_before = False
                if i > 0 and parts:
                    prev_part = parts[-1].strip() if parts else ""
                    current_part = part.strip()

                    if (
                        child_tag == "mo"
                        and current_part
                        in [
                            "+",
                            "-",
                            "=",
                            "<",
                            ">",
                            "\\times",
                            "\\div",
                            "\\pm",
                            "\\leq",
                            "\\geq",
                            "\\neq",
                            "\\approx",
                        ]
                        and prev_part
                        and not prev_part.endswith(" ")
                    ):
                        need_space_before = True

                    if i > 0:
                        prev_child = children[i - 1]
                        prev_tag = etree.QName(prev_child).localname
                        if prev_tag == "mo" and child_tag != "mo":
                            prev_text = prev_child.text or ""
                            if (
                                prev_text.strip() in ["+", "-", "=", "<", ">", "×", "÷", "±", "≤", "≥", "≠", "≈"]
                                and prev_part
                                and not prev_part.endswith(" ")
                            ):
                                need_space_before = True

                if need_space_before:
                    parts.append(" ")
                    logger.debug("Added space around operator")

                parts.append(part)

            i += 1

        return _join_latex(parts)

    elif tag == "mtable":
        matrix_content = _convert_matrix_rows(node)
        return f"\\begin{{matrix}}\n{matrix_content}\n\\end{{matrix}}"

    elif tag == "mfenced":
        open_char = node.get("open", "(")
        close_char = node.get("close", ")")

        matrix = _single_matrix(node[0]) if len(node) == 1 else None
        matrix_type = {
            ("[", "]"): "bmatrix",
            ("(", ")"): "pmatrix",
            ("|", "|"): "vmatrix",
            ("‖", "‖"): "Vmatrix",
            ("||", "||"): "Vmatrix",
            ("{", "}"): "Bmatrix",
        }.get((open_char, close_char))
        if matrix is not None and matrix_type is not None:
            matrix_content = _convert_matrix_rows(matrix)
            return f"\\begin{{{matrix_type}}}\n{matrix_content}\n\\end{{{matrix_type}}}"
        parts = [_mathml_node_to_latex(child) for child in node]
        separators = node.get("separators", ",").replace(" ", "")
        joined: list[str] = []
        for index, part in enumerate(parts):
            if index and separators:
                joined.append(_math_literal(separators[min(index - 1, len(separators) - 1)]))
            joined.append(part)
        content = _join_latex(joined)
        if not open_char and not close_char:
            return "{" + content + "}"
        return f"{_math_literal(open_char)}{content}{_math_literal(close_char)}"

    else:
        raise ValueError(f"unsupported MathML structure: {tag}")


def mathml_to_latex(mathml_str: str) -> str | None:
    """Convert MathML to LaTeX format.

    Args:
        mathml_str: MathML XML string.

    Returns:
        LaTeX string, or None if conversion fails.
    """
    try:
        mathml_tree = etree.fromstring(mathml_str.encode("utf-8"))
        return _mathml_node_to_latex(mathml_tree)
    except Exception as e:
        logger.error("MathML to LaTeX failed: %s", e)
        return None
