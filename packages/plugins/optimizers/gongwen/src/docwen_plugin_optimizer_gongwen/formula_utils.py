"""OMML formula detection and LaTeX conversion for gongwen paragraphs.

Delegates to ``docwen_core.formula`` for the OMML→MathML→LaTeX chain
and the docx plugin's ``formula_extractor`` for the thin adapter.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph

from docwen_core.docx_parsing.xml_ns import NS_M, NS_W

logger = logging.getLogger(__name__)


def extract_formula_info(para: Paragraph) -> tuple[bool, str, str]:
    """Extract formula detection info from a paragraph.

    Uses the existing ``docwen_core.formula`` OMML→MathML→LaTeX
    conversion chain (``omml_to_mathml`` + ``mathml_to_latex``) and the
    docx plugin's ``formula_extractor.extract_formula_from_element``.

    Args:
        para: A python-docx Paragraph.

    Returns:
        (has_formula, formula_type, formula_latex) tuple.
        formula_type is ``"block"`` for standalone formula paragraphs,
        ``"inline"`` for mixed text+formula paragraphs, or ``""`` if no
        formula was found.
        formula_latex is the LaTeX representation (e.g. ``\\sum_{i=1}^n x_i``).
    """
    nsmap = {"m": NS_M, "w": NS_W}

    try:
        omaths = para._p.findall(".//m:oMath", nsmap)
        omath_paras = para._p.findall(".//m:oMathPara", nsmap)
        all_omaths = list(omaths) + list(omath_paras)

        if not all_omaths:
            return False, "", ""

        # Determine inline vs block
        para_text = para.text.strip() if para.text else ""
        text_length = len(para_text)

        # If paragraph only has formula content or very short text, treat as block
        formula_type = "block" if text_length < 20 else "inline"

        # Convert each OMML element to LaTeX via the shared chain
        latex_parts: list[str] = []
        for om_elem in all_omaths:
            latex = _omml_element_to_latex(om_elem)
            if latex:
                latex_parts.append(latex)

        formula_latex = " ".join(latex_parts) if latex_parts else ""

        return True, formula_type, formula_latex

    except Exception:
        logger.debug("Error extracting formula info", exc_info=True)
        return False, "", ""


def _omml_element_to_latex(om_elem) -> str:
    """Convert a single OMML XML element to LaTeX via the core formula chain.

    Uses ``omml_to_mathml`` (pure-Python recursive conversion) and
    ``mathml_to_latex`` (tree-walking LaTeX generator) from
    ``docwen_core.formula``.
    """
    import lxml.etree as etree

    try:
        from docwen_core.formula import mathml_to_latex, omml_to_mathml

        omml_xml = etree.tostring(om_elem, encoding="unicode")
        mathml = omml_to_mathml(omml_xml)
        if not mathml:
            return ""
        return mathml_to_latex(mathml) or ""
    except Exception:
        logger.debug("OMML→LaTeX conversion failed", exc_info=True)
        return ""
