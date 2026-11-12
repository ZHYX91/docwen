"""Word SDT boundary formatting cannot carry or relax semantic ownership."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from lxml import etree

from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error
from docwen_core._docx_semantics_v3_topology import (
    prove_block_sdt_envelope,
    prove_inline_occurrence_envelope,
    prove_soft_reference_envelope,
)
from docwen_core.docx_host_metadata import owned_sdt_parts
from docwen_core.docx_numbering_occurrence import (
    derive_numbering_occurrence,
    prove_numbering_occurrence_sdt,
    wrap_numbering_occurrence,
)
from docwen_core.docx_standalone_caption_occurrence import (
    derive_standalone_caption_occurrence,
    prove_standalone_caption_occurrence_sdt,
    wrap_standalone_caption_occurrence,
)

pytestmark = pytest.mark.contract

_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_WORD_DEFAULTS = (
    '<w:rPr><w:b/><w:bCs/><w:color w:val="4F81BD" w:themeColor="accent1"/>'
    '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
)
_WORD_END_DEFAULTS = (
    '<w:rPr><w:b w:val="0"/><w:bCs w:val="0"/><w:color w:val="auto"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
)


def _control(*, end_properties: bool = True, end_defaults: str = "") -> Any:
    end = f"<w:sdtEndPr>{end_defaults}</w:sdtEndPr>" if end_properties else ""
    return parse_xml(
        f'<w:sdt {_NS}><w:sdtPr>{_WORD_DEFAULTS}<w:tag w:val="owned"/>'
        f'<w:id w:val="-42"/></w:sdtPr>{end}'
        "<w:sdtContent><w:r><w:t>token</w:t></w:r></w:sdtContent></w:sdt>"
    )


def _prove(sdt: Any, family: str) -> None:
    if family == "block":
        prove_block_sdt_envelope(sdt, "owned")
    elif family == "inline":
        prove_inline_occurrence_envelope(sdt, "owned")
    else:
        prove_soft_reference_envelope(sdt, "owned", "token")


def _damage(sdt: Any, damage: str) -> None:
    properties, end, content = list(sdt)
    if damage == "wrong_root":
        sdt.tag = qn("w:p")
    elif damage == "duplicate_end":
        sdt.insert(1, deepcopy(end))
    elif damage == "duplicate_properties":
        sdt.insert(0, deepcopy(properties))
    elif damage == "duplicate_content":
        sdt.append(deepcopy(content))
    elif damage == "end_before_properties":
        sdt.remove(end)
        sdt.insert(0, end)
    elif damage == "end_after_content":
        sdt.remove(end)
        sdt.append(end)
    elif damage == "end_child":
        end.append(OxmlElement("w:r"))
    elif damage == "extra_run":
        sdt.append(OxmlElement("w:r"))
    elif damage == "missing_properties":
        sdt.remove(properties)
    elif damage == "missing_content":
        sdt.remove(content)
    else:
        owner, payload = damage.split("_", 1)
        node = {"outer": sdt, "properties": properties, "content": content, "end": end}[owner]
        if payload == "attribute":
            node.set(qn("w:val"), "injected")
        else:
            setattr(node, payload, "injected")


@pytest.mark.parametrize(
    ("end_properties", "end_defaults"),
    [(False, ""), (True, ""), (True, _WORD_END_DEFAULTS), (True, _WORD_END_DEFAULTS + _WORD_DEFAULTS)],
)
def test_sdt_parts_returns_original_nodes_without_normalizing_content(end_properties: bool, end_defaults: str) -> None:
    sdt = _control(end_properties=end_properties, end_defaults=end_defaults)
    original = etree.tostring(sdt)

    parts = owned_sdt_parts(sdt)

    assert parts is not None
    assert parts[0] is sdt.find(qn("w:sdtPr"))
    assert parts[1] is sdt.find(qn("w:sdtContent"))
    assert etree.tostring(sdt) == original


@pytest.mark.parametrize(
    "damage",
    [
        "wrong_root",
        "outer_attribute",
        "outer_text",
        "outer_tail",
        "properties_attribute",
        "properties_text",
        "properties_tail",
        "content_attribute",
        "content_text",
        "content_tail",
        "end_attribute",
        "end_text",
        "end_tail",
        "end_child",
        "duplicate_end",
        "duplicate_properties",
        "duplicate_content",
        "end_before_properties",
        "end_after_content",
        "extra_run",
        "missing_properties",
        "missing_content",
    ],
)
def test_sdt_parts_rejects_extra_payload_and_ambiguous_order(damage: str) -> None:
    sdt = _control()
    _damage(sdt, damage)

    assert owned_sdt_parts(sdt) is None


@pytest.mark.parametrize("family", ["block", "inline", "soft"])
@pytest.mark.parametrize(
    ("end_properties", "end_defaults"),
    [(False, ""), (True, ""), (True, _WORD_END_DEFAULTS), (True, _WORD_END_DEFAULTS + _WORD_DEFAULTS)],
)
def test_owned_envelopes_accept_validated_host_end_properties(
    family: str, end_properties: bool, end_defaults: str
) -> None:
    sdt = _control(end_properties=end_properties, end_defaults=end_defaults)
    original = etree.tostring(sdt)

    _prove(sdt, family)

    assert etree.tostring(sdt) == original


@pytest.mark.parametrize("family", ["block", "inline", "soft"])
@pytest.mark.parametrize(
    "damage",
    ["end_attribute", "end_text", "end_tail", "end_child", "duplicate_end", "end_after_content"],
)
def test_host_metadata_does_not_admit_a_modified_envelope(family: str, damage: str) -> None:
    sdt = _control()
    _damage(sdt, damage)

    with pytest.raises(DocxSemanticsV3Error, match="canonical"):
        _prove(sdt, family)


@pytest.mark.parametrize(
    "defaults",
    [
        "<w:rPr><w:t>hidden text</w:t></w:rPr>",
        '<w:rPr><w:fldChar w:fldCharType="begin"/></w:rPr>',
        "<w:rPr><w:drawing/></w:rPr>",
        "<w:rPr><w:rPrChange/></w:rPr>",
        "<w:rPr><w:b/><w:b/></w:rPr>",
        '<w:rPr><w:b w:val="perhaps"/></w:rPr>',
        '<w:rPr><w:color w:val="invalid"/></w:rPr>',
        '<w:rPr><w:sz w:val="-22"/></w:rPr>',
        '<w:rPr w:unknown="1"/>',
        "<w:rPr>hidden text</w:rPr>",
        "<w:rPr><w:b/>hidden tail</w:rPr>",
        "<w:rPr/>hidden tail",
    ],
)
def test_end_boundary_defaults_reject_invalid_values_and_nonvisual_payload(defaults: str) -> None:
    sdt = _control(end_defaults=defaults)
    assert owned_sdt_parts(sdt) is None
    with pytest.raises(DocxSemanticsV3Error, match="canonical"):
        prove_block_sdt_envelope(sdt, "owned")


@pytest.mark.parametrize("family", ["block", "inline", "soft"])
@pytest.mark.parametrize("damage", ["wrong_tag", "duplicate_tag"])
def test_empty_end_properties_do_not_replace_exact_tag_authority(family: str, damage: str) -> None:
    sdt = _control()
    properties = sdt.find(qn("w:sdtPr"))
    tag = properties.find(qn("w:tag"))
    if damage == "wrong_tag":
        tag.set(qn("w:val"), "other")
    else:
        properties.append(deepcopy(tag))
    assert owned_sdt_parts(sdt) is not None

    with pytest.raises(DocxSemanticsV3Error, match="canonical"):
        _prove(sdt, family)


def test_soft_reference_content_still_has_one_exact_visible_run() -> None:
    sdt = _control()
    content = sdt.find(qn("w:sdtContent"))
    content.append(deepcopy(content[0]))
    assert owned_sdt_parts(sdt) is not None

    with pytest.raises(DocxSemanticsV3Error, match="canonical"):
        prove_soft_reference_envelope(sdt, "owned", "token")


def _caption_occurrence(paired: bool, *, end_defaults: str = "") -> tuple[Any, Any, Any]:
    document = Document()
    caption = document.add_paragraph("Results", style="Caption")
    parameters: dict[str, Any] = {
        "source_sha256": "a" * 64,
        "source_start": 0,
        "source_end": 12,
        "kind": "table",
        "plan_sha256": "b" * 64,
    }
    identity: Any
    if paired:
        identity = derive_numbering_occurrence(**parameters)
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "value"
        wrap_numbering_occurrence(caption._p, (table._tbl,), identity)
    else:
        identity = derive_standalone_caption_occurrence(
            **parameters, enabled=False, target_id=None, derived_number=None
        )
        wrap_standalone_caption_occurrence(caption._p, identity)
    content = caption._p.getparent()
    assert content is not None
    wrapper = content.getparent()
    assert wrapper is not None
    wrapper.insert(1, parse_xml(f"<w:sdtEndPr {_NS}>{end_defaults}</w:sdtEndPr>"))
    return document, wrapper, identity


def _prove_caption_occurrence(sdt: Any, identity: Any, paired: bool) -> None:
    if paired:
        prove_numbering_occurrence_sdt(sdt, identity, caption_style_id="Caption")
    else:
        prove_standalone_caption_occurrence_sdt(sdt, identity, caption_style_id="Caption")


@pytest.mark.parametrize("paired", [False, True])
@pytest.mark.parametrize("end_defaults", ["", _WORD_END_DEFAULTS])
def test_caption_occurrence_keeps_exact_slots_after_saved_end_properties(
    tmp_path: Path, paired: bool, end_defaults: str
) -> None:
    document, _wrapper, identity = _caption_occurrence(paired, end_defaults=end_defaults)
    output = tmp_path / "caption-occurrence.docx"
    document.save(str(output))
    reopened = Document(str(output))
    [sdt] = list(reopened.element.iter(qn("w:sdt")))

    _prove_caption_occurrence(sdt, identity, paired)


@pytest.mark.parametrize("paired", [False, True])
@pytest.mark.parametrize("damage", ["duplicate_tag", "end_child", "extra_content"])
def test_caption_occurrence_end_metadata_does_not_hide_extra_identity_or_blocks(paired: bool, damage: str) -> None:
    _document, sdt, identity = _caption_occurrence(paired)
    if damage == "duplicate_tag":
        properties = sdt.find(qn("w:sdtPr"))
        properties.append(deepcopy(properties.find(qn("w:tag"))))
    elif damage == "end_child":
        sdt.find(qn("w:sdtEndPr")).append(OxmlElement("w:r"))
    else:
        sdt.find(qn("w:sdtContent")).append(OxmlElement("w:p"))

    with pytest.raises(DocxSemanticsV3Error):
        _prove_caption_occurrence(sdt, identity, paired)
