"""Thin adapter: DOCX OMML elements → Markdown LaTeX strings."""

from __future__ import annotations

import logging
from collections.abc import Callable, Mapping

from docwen_core.conversion_loss import ConversionLoss, loss_text
from docwen_core.formula import mathml_to_latex, omml_to_mathml
from docwen_core.markdown_math import serialize_formula

logger = logging.getLogger(__name__)


def omml_to_markdown(
    omml_xml: str,
    *,
    block: bool,
    compact: bool = False,
    messages: Mapping[str, str] | None = None,
    on_failure: Callable[[ConversionLoss], None] | None = None,
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
    loss = ConversionLoss(
        code="DOCX2MD-FORMULA-SOURCE-FALLBACK",
        reason=loss_text(messages, "formula_reason", reason),
        change=loss_text(
            messages,
            "formula_change",
            "The formula could not be converted; a visible placeholder was retained. Check the source formula.",
        ),
    )
    if on_failure is not None:
        on_failure(loss)
    else:
        logger.warning(loss.message())
    return loss_text(messages, "formula_placeholder", "[Formula could not be converted]")


def extract_formula_from_element(
    element, *, messages: Mapping[str, str] | None = None, on_failure: Callable[[ConversionLoss], None] | None = None
) -> str:
    """Convert a DOCX XML math element to Markdown formula text."""
    import lxml.etree as etree

    xml = etree.tostring(element, encoding="unicode")
    block = etree.QName(element).localname == "oMathPara"
    in_cell = any(etree.QName(ancestor).localname == "tc" for ancestor in element.iterancestors())
    return omml_to_markdown(xml, block=block, compact=in_cell, messages=messages, on_failure=on_failure)
