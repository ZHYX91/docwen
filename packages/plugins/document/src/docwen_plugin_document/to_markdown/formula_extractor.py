"""Thin adapter: DOCX OMML elements → Markdown LaTeX strings."""

from __future__ import annotations

from docwen_core.formula import mathml_to_latex, omml_to_mathml


def omml_to_markdown(omml_xml: str, *, block: bool) -> str:
    """Convert OMML XML to ``$...$`` or ``$$...$$`` Markdown."""
    if not omml_xml.strip():
        return ""
    mathml = omml_to_mathml(omml_xml)
    if not mathml:
        return ""
    latex = mathml_to_latex(mathml)
    if not latex:
        return ""
    return f"$${latex}$$" if block else f"${latex}$"


def extract_formula_from_element(element, *, block: bool) -> str:
    """Convert a DOCX XML math element to Markdown formula text."""
    import lxml.etree as etree

    xml = etree.tostring(element, encoding="unicode")
    return omml_to_markdown(xml, block=block)
