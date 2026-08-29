"""Private caption-carrier evidence stays independent from semantic labels."""

from __future__ import annotations

import pytest
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_core._docx_caption_carrier import captionable_logical_elements
from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("carrier", ["table", "drawing", "equation", "fenced_code"])
def test_private_caption_carrier_evidence_accepts_all_native_object_families(carrier: str) -> None:
    element = _carrier(carrier)

    assert captionable_logical_elements((element,)) == (element,)


def test_private_caption_carrier_evidence_rejects_an_ordinary_paragraph() -> None:
    paragraph = OxmlElement("w:p")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "ordinary prose"
    run.append(text)
    paragraph.append(run)

    with pytest.raises(DocxSemanticsV3Error, match="lacks image, equation, or fenced-code evidence"):
        captionable_logical_elements((paragraph,))


def _carrier(kind: str):
    if kind == "table":
        return OxmlElement("w:tbl")
    paragraph = OxmlElement("w:p")
    if kind == "drawing":
        run = OxmlElement("w:r")
        run.append(OxmlElement("w:drawing"))
        paragraph.append(run)
        return paragraph
    if kind == "equation":
        paragraph.append(OxmlElement("m:oMath"))
        return paragraph
    if kind == "fenced_code":
        sdt = OxmlElement("w:sdt")
        properties = OxmlElement("w:sdtPr")
        tag = OxmlElement("w:tag")
        tag.set(qn("w:val"), "docwen-fenced-source-v1:" + "a" * 32)
        properties.append(tag)
        sdt.append(properties)
        sdt.append(OxmlElement("w:sdtContent"))
        paragraph.append(sdt)
        return paragraph
    raise AssertionError(kind)
