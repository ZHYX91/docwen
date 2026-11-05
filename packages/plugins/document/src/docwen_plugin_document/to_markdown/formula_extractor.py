"""Thin adapter: DOCX OMML elements → Markdown LaTeX strings."""

from __future__ import annotations

import json
import logging
import re
from collections.abc import Callable

from docwen_core.formula import mathml_to_latex, omml_to_mathml
from docwen_core.markdown_math import serialize_formula

logger = logging.getLogger(__name__)


def omml_to_markdown(
    omml_xml: str,
    *,
    block: bool,
    compact: bool = False,
    on_failure: Callable[[str], None] | None = None,
) -> str:
    """Convert OMML XML to ``$...$`` or ``$$...$$`` Markdown."""
    if not omml_xml.strip():
        return ""
    mathml = omml_to_mathml(omml_xml)
    latex = mathml_to_latex(mathml) if mathml else None
    if latex:
        return serialize_formula(latex, block=block, compact=compact)
    reason = (
        "OMML contains an unsupported or invalid structure" if not mathml else "MathML cannot be represented as LaTeX"
    )
    message = f"{reason}; the complete original OMML was preserved as a JSON string in a code span."
    if on_failure is not None:
        on_failure(message)
    else:
        logger.warning(message)
    # JSON retains original newlines and quotes inside a single physical
    # Markdown line, including in tables and list items. Nothing is silently
    # discarded, and no invented LaTeX is presented as the original formula.
    literal = "OMML " + json.dumps(omml_xml, ensure_ascii=False)
    marker = "`" * (max((len(item) for item in re.findall(r"`+", literal)), default=0) + 1)
    return f"{marker}{literal}{marker}"


def extract_formula_from_element(element, *, on_failure: Callable[[str], None] | None = None) -> str:
    """Convert a DOCX XML math element to Markdown formula text."""
    import lxml.etree as etree

    xml = etree.tostring(element, encoding="unicode")
    block = etree.QName(element).localname == "oMathPara"
    in_cell = any(etree.QName(ancestor).localname == "tc" for ancestor in element.iterancestors())
    return omml_to_markdown(xml, block=block, compact=in_cell, on_failure=on_failure)
