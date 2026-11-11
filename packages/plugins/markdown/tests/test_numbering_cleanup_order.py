"""Every writer preserves the trailing cleanup element in valid numbering XML."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from tests.support.numbering import repository_numbering_registry

from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from ._resolved_v4_support import _NEUTRAL, _PLAN, _context, _refs
from .conftest import make_context

pytestmark = pytest.mark.contract

# Independent schema excerpt: validates sequence only, not a complete DOCX.
_SCHEMA = etree.XMLSchema(
    etree.fromstring(b"""<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"
  targetNamespace="http://schemas.openxmlformats.org/wordprocessingml/2006/main" elementFormDefault="qualified">
  <xs:element name="numbering"><xs:complexType><xs:sequence>
    <xs:element name="numPicBullet" type="xs:anyType" minOccurs="0" maxOccurs="unbounded"/>
    <xs:element name="abstractNum" type="xs:anyType" minOccurs="0" maxOccurs="unbounded"/>
    <xs:element name="num" type="xs:anyType" minOccurs="0" maxOccurs="unbounded"/>
    <xs:element name="numIdMacAtCleanup" type="xs:anyType" minOccurs="0"/>
  </xs:sequence><xs:anyAttribute processContents="lax"/></xs:complexType></xs:element>
</xs:schema>""")
)


@pytest.mark.parametrize("existing", [False, True])
@pytest.mark.parametrize("cleanup", [False, True])
@pytest.mark.parametrize("route", ["list", "heading", "resolved", "resolved_list"])
def test_saved_numbering_has_schema_order_and_keeps_template_definitions(
    tmp_path: Path, existing: bool, cleanup: bool, route: str
) -> None:
    document = Document()
    document.add_paragraph("{{body}}")
    numbering = document.part.numbering_part.element
    if not existing:
        for child in list(numbering):
            numbering.remove(child)
    if cleanup:
        marker = OxmlElement("w:numIdMacAtCleanup")
        marker.set(qn("w:val"), "12")
        numbering.append(marker)
    _SCHEMA.assertValid(numbering)
    before = [
        (child.tag, dict(child.attrib), etree.tostring(child, method="c14n", exclusive=True)) for child in numbering
    ]
    template = tmp_path / "template.docx"
    document.save(template)
    original = template.read_bytes()
    options = {"template_name": str(template), "heading_merge_mode": "never"}
    if route.startswith("resolved"):
        neutral_path, plan_path = _NEUTRAL, _PLAN
        if route == "resolved_list":
            neutral = json.loads(_NEUTRAL.read_text(encoding="utf-8"))
            plan = json.loads(_PLAN.read_text(encoding="utf-8"))
            neutral["document"]["authored_markdown"] += "\n\n1. New item\n2. Next item\n"
            digest = hashlib.sha256(neutral["document"]["authored_markdown"].encode()).hexdigest()
            neutral["source_sha256"] = plan["source_sha256"] = digest
            neutral_path, plan_path = tmp_path / "neutral.json", tmp_path / "plan.json"
            neutral_path.write_text(json.dumps(neutral), encoding="utf-8")
            plan_path.write_text(json.dumps(plan), encoding="utf-8")
        context, _ = _context(tmp_path, _refs(neutral_path, plan_path), options=options)
    else:
        source = tmp_path / "source.md"
        source.write_text("# Heading\n" if route == "heading" else "1. New item\n2. Next item\n", encoding="utf-8")
        if route == "heading":
            options.update(
                add_numbering=True,
                heading_numbering_render_mode="word_native",
                numbering_scheme="hierarchical_standard",
            )
        context, _ = make_context(str(source), options=options, numbering_registry=repository_numbering_registry())
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    output = Document(result.artifacts[0].staging_path).part.numbering_part.element
    _SCHEMA.assertValid(output)
    for tag, attributes, payload in before:
        retained = [child for child in output if child.tag == tag and dict(child.attrib) == attributes]
        assert len(retained) == 1
        assert etree.tostring(retained[0], method="c14n", exclusive=True) == payload
    if cleanup:
        assert output[-1].tag == qn("w:numIdMacAtCleanup")
    assert template.read_bytes() == original
