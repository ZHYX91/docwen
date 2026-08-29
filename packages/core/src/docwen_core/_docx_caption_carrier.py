"""Private structural proof for captionable DOCX carrier objects."""

from __future__ import annotations

from typing import Any

from docwen_core._docx_semantics_v3_fenced import FENCED_SOURCE_TAG_PREFIX
from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error
from docwen_core._docx_semantics_v3_ooxml import sdt_tag
from docwen_core._docx_semantics_v3_topology import logical_group_elements

_MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def captionable_logical_elements(elements: tuple[Any, ...]) -> tuple[Any, ...]:
    """Return a proven native table, image, equation, or fenced-code carrier.

    Semantic caption kind is intentionally absent: a Figure may describe a
    native table composite.  Paragraph carriers still require private physical
    evidence, so an ordinary paragraph cannot become a caption object merely
    because it was placed in a target wrapper.
    """

    from docx.oxml.ns import qn

    logical = logical_group_elements(elements)
    if len(logical) == 1 and logical[0].tag == qn("w:tbl"):
        return logical
    if not logical or any(item.tag != qn("w:p") for item in logical):
        raise DocxSemanticsV3Error("caption target has no supported captionable object")

    if len(logical) == 1 and (
        logical[0].find(f".//{qn('w:drawing')}") is not None
        or logical[0].find(f".//{qn('w:pict')}") is not None
        or logical[0].find(f".//{{{_MATH_NAMESPACE}}}oMath") is not None
        or logical[0].find(f".//{{{_MATH_NAMESPACE}}}oMathPara") is not None
    ):
        return logical
    if all(_has_fenced_source_carrier(item) for item in logical):
        return logical
    raise DocxSemanticsV3Error("caption paragraph carrier lacks image, equation, or fenced-code evidence")


def _has_fenced_source_carrier(paragraph: Any) -> bool:
    from docx.oxml.ns import qn

    return any((sdt_tag(item) or "").startswith(FENCED_SOURCE_TAG_PREFIX) for item in paragraph.iter(qn("w:sdt")))


__all__ = ["captionable_logical_elements"]
