"""Actual conversion and fresh proof retain relocated and unrelated template parts."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.opc.packuri import PackURI
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from lxml import etree
from tests.support.numbering import repository_numbering_registry

from docwen_core.docx_numbering_import import DocxNumberingImportError, HeadingNumberingProofIndex
from docwen_core.docx_package_parts import NUMBERING, STYLES, STYLES_EFFECTS
from docwen_core.docx_resolved_numbering_recovery import ResolvedNumberingV4Recovery
from docwen_core.docx_semantics_v3 import DocxSemanticsV3Recovery
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from ._resolved_v4_support import _NEUTRAL, _PLAN, _context, _refs
from .conftest import make_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("route", ["ordinary", "resolved"])
@pytest.mark.parametrize("moved", ["styles", "numbering", "both"])
def test_relocated_registries_pass_writing_and_fresh_recovery(tmp_path: Path, route: str, moved: str) -> None:
    document = Document()
    document.add_paragraph("{{body}}")
    relocated = []
    for role in (STYLES, STYLES_EFFECTS, NUMBERING):
        if (role is NUMBERING and moved == "styles") or (role is not NUMBERING and moved == "numbering"):
            continue
        part = document.part.part_related_by(role.relationship_type)
        name = f"/registries/nested/{role.name}.xml"
        part._partname = PackURI(name)
        relocated.append((role, name))
    template = tmp_path / "template.docx"
    document.save(str(template))
    original = template.read_bytes()
    options: dict[str, object] = {"template_name": str(template), "heading_merge_mode": "never"}
    if route == "resolved":
        context, _ = _context(tmp_path, _refs(_NEUTRAL, _PLAN), options=options)
    else:
        source = tmp_path / "input.md"
        source.write_text(
            "# Generated\n\n1. Item\n\nTable: Results ^results\n\n| A |\n|---|\n| 1 |\n", encoding="utf-8"
        )
        options.update(
            markdown_extensions={"input": {"captions_references": True}},
            add_numbering=True,
            heading_numbering_render_mode="word_native",
            numbering_scheme="hierarchical_standard",
        )
        context, _ = make_context(str(source), options=options, numbering_registry=repository_numbering_registry())
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    output = Path(result.artifacts[0].staging_path)
    reopened = Document(str(output))
    for role, name in relocated:
        assert str(reopened.part.part_related_by(role.relationship_type).partname) == name
    with ZipFile(template) as before, ZipFile(output) as after:
        prior = etree.fromstring(before.read("word/_rels/document.xml.rels"))
        final = {
            item.get("Id"): dict(item.attrib) for item in etree.fromstring(after.read("word/_rels/document.xml.rels"))
        }
        assert all(final[item.get("Id")] == dict(item.attrib) for item in prior)
        assert after.read("[Content_Types].xml").count(b"/registries/nested/") == len(relocated)
    if route == "resolved":
        recovery = ResolvedNumberingV4Recovery.load_if_present(output, reopened)
        assert recovery is not None and recovery.caption_signatures
    else:
        recovery = DocxSemanticsV3Recovery.load(output, reopened)
        assert recovery.caption_signatures == (("table", "results", "Results", "1"),)
    assert template.read_bytes() == original


def _full_override(numbering: etree._Element, instance: etree._Element) -> etree._Element:
    reference = instance.find(qn("w:abstractNumId"))
    assert reference is not None
    abstract = next(
        node
        for node in numbering.findall(qn("w:abstractNum"))
        if node.get(qn("w:abstractNumId")) == reference.get(qn("w:val"))
    )
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    template_level = abstract.find(qn("w:lvl"))
    assert template_level is not None
    level = deepcopy(template_level)
    number_format = level.find(qn("w:numFmt"))
    assert number_format is not None
    number_format.set(qn("w:val"), "upperRoman")
    override.append(level)
    instance.append(override)
    return override


@pytest.mark.parametrize("template_reference", [False, True])
def test_unrelated_full_override_survives_while_owned_heading_proof_stays_strict(
    tmp_path: Path, template_reference: bool
) -> None:
    document = Document()
    numbering = document.part.numbering_part.element
    old_instance = numbering.find(qn("w:num"))
    assert old_instance is not None
    old_id = old_instance.get(qn("w:numId"))
    assert old_id is not None
    _full_override(numbering, old_instance)
    snapshot = etree.tostring(old_instance, method="c14n", exclusive=True)
    if template_reference:
        fixed = document.add_paragraph("Fixed template item")
        num_pr = OxmlElement("w:numPr")
        level = OxmlElement("w:ilvl")
        level.set(qn("w:val"), "0")
        num_id = OxmlElement("w:numId")
        num_id.set(qn("w:val"), old_id)
        num_pr.extend((level, num_id))
        fixed._p.get_or_add_pPr().append(num_pr)
    document.add_paragraph("{{body}}")
    template = tmp_path / "template.docx"
    document.save(str(template))
    original = template.read_bytes()
    context, _ = _context(
        tmp_path, _refs(_NEUTRAL, _PLAN), options={"template_name": str(template), "heading_merge_mode": "never"}
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    output = Path(result.artifacts[0].staging_path)
    reopened = Document(str(output))
    output_numbering = reopened.part.numbering_part.element
    retained = next(node for node in output_numbering.findall(qn("w:num")) if node.get(qn("w:numId")) == old_id)
    assert etree.tostring(retained, method="c14n", exclusive=True) == snapshot
    assert ResolvedNumberingV4Recovery.load_if_present(output, reopened) is not None
    assert template.read_bytes() == original

    body = reopened.element.find(qn("w:body"))
    assert body is not None
    generated = next(
        element
        for element in body.iter(qn("w:p"))
        if (node := element.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")) is not None
        and node.get(qn("w:val")) not in {old_id, "0"}
    )
    assert isinstance(generated, CT_P)
    generated_reference = generated.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")
    assert generated_reference is not None
    generated_id = generated_reference.get(qn("w:val"))
    generated_instance = next(
        node for node in output_numbering.findall(qn("w:num")) if node.get(qn("w:numId")) == generated_id
    )
    _full_override(output_numbering, generated_instance)
    damaged = tmp_path / "owned-override.docx"
    reopened.save(str(damaged))
    index = HeadingNumberingProofIndex.load(damaged)
    with pytest.raises(DocxNumberingImportError, match="exact ancestry"):
        index.prove_paragraph(Paragraph(generated, reopened))
