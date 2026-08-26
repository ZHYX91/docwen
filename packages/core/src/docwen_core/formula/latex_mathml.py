"""LaTeX to MathML conversion, with dedicated matrix support."""

from __future__ import annotations

import logging
import re

import lxml.etree as etree

from docwen_core.formula.constants import (
    LATEX2MATHML_AVAILABLE,
    MATHML_NS,
    MATRIX_BRACKET_MAP,
    _latex_to_mathml_convert,
)

logger = logging.getLogger(__name__)


def _convert_cell_content_to_mathml(cell_str: str, parent_elem) -> None:
    """Convert matrix cell content to MathML nodes.

    Supports: simple numbers, simple variables, subscripts (a_{11}),
    superscripts (x^2), Greek letters (\\alpha).
    """
    cell_str = cell_str.strip()

    if not cell_str:
        return

    # Try latex2mathml for cells with complex content
    if LATEX2MATHML_AVAILABLE and ("_" in cell_str or "^" in cell_str or "\\" in cell_str):
        try:
            mathml_str = _latex_to_mathml_convert(cell_str)
            if mathml_str:
                temp_tree = etree.fromstring(mathml_str.encode("utf-8"))
                for child in temp_tree:
                    parent_elem.append(child)
                return
        except Exception as e:
            logger.debug(f"Cell latex2mathml conversion failed, using simple processing: {e}")

    # Simple number
    if cell_str.isdigit() or (cell_str.startswith("-") and cell_str[1:].isdigit()):
        mn = etree.SubElement(parent_elem, "mn")
        mn.text = cell_str
    else:
        # Simple identifier
        mi = etree.SubElement(parent_elem, "mi")
        mi.text = cell_str


def _convert_matrix_latex_to_mathml(latex_str: str) -> str | None:
    """Manually convert a single LaTeX matrix to MathML.

    Supports: bmatrix, pmatrix, matrix, vmatrix, Vmatrix, Bmatrix.
    Only handles a single matrix in the input.
    """
    try:
        matrix_count = len(re.findall(r"\\begin\{([bpBvV]?matrix)\}", latex_str))
        if matrix_count > 1:
            logger.debug(f"Detected {matrix_count} matrices, skipping manual conversion")
            return None

        pattern = r"\\begin\{([bpBvV]?matrix)\}(.+?)\\end\{\1\}"
        match = re.search(pattern, latex_str, re.DOTALL)

        if not match:
            return None

        matrix_type = match.group(1)
        matrix_content = match.group(2).strip()

        # Only convert if the matrix is the entire formula
        if match.group(0).strip() != latex_str.strip():
            logger.debug("LaTeX contains content outside the matrix, skipping manual conversion")
            return None

        rows = [row.strip() for row in matrix_content.split(r"\\") if row.strip()]

        math = etree.Element("math", xmlns=MATHML_NS)

        open_bracket, close_bracket = MATRIX_BRACKET_MAP.get(matrix_type, ("", ""))

        if open_bracket and close_bracket:
            mfenced = etree.SubElement(math, "mfenced")
            mfenced.set("open", open_bracket)
            mfenced.set("close", close_bracket)
            table = etree.SubElement(mfenced, "mtable")
        else:
            table = etree.SubElement(math, "mtable")

        for row_str in rows:
            mtr = etree.SubElement(table, "mtr")
            cells = [cell.strip() for cell in row_str.split("&")]
            for cell_str in cells:
                mtd = etree.SubElement(mtr, "mtd")
                if cell_str:
                    _convert_cell_content_to_mathml(cell_str, mtd)

        logger.info(f"Manual single matrix conversion succeeded: {matrix_type}")
        return etree.tostring(math, encoding="unicode")
    except Exception as e:
        logger.error(f"Manual matrix conversion failed: {e}")
        return None


def latex_to_mathml(latex_str: str) -> str | None:
    """Convert LaTeX formula to MathML format.

    Args:
        latex_str: Raw LaTeX string.

    Returns:
        MathML XML string, or None if conversion fails.
    """
    if not LATEX2MATHML_AVAILABLE:
        logger.error("latex2mathml library not installed")
        return None

    try:
        latex_str = latex_str.strip()
        if not latex_str:
            return None

        # Try manual matrix conversion first for matrix environments
        if r"\begin{" in latex_str and "matrix" in latex_str:
            logger.debug("Matrix environment detected, trying manual conversion")
            mathml = _convert_matrix_latex_to_mathml(latex_str)
            if mathml:
                return mathml

        # Fall back to latex2mathml library
        mathml = _latex_to_mathml_convert(latex_str)
        return mathml
    except Exception as e:
        logger.error(f"LaTeX to MathML failed: {e}, LaTeX: {latex_str}")
        # Second attempt: try manual matrix conversion on failure
        if r"\begin{" in latex_str and "matrix" in latex_str:
            logger.debug("latex2mathml failed, trying manual matrix conversion")
            return _convert_matrix_latex_to_mathml(latex_str)
        return None
