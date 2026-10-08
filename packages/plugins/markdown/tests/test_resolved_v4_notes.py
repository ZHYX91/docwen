"""Resolved-document note warnings follow actual template numbering."""

import json
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from .test_resolved_v4_to_docx import _NEUTRAL, _PLAN, _context, _refs, _sha_text

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("restart,expected", [("eachPage", "?"), ("continuous", "1")])
def test_resolved_repeated_notes_report_deferred_template_numbering(
    tmp_path: Path, restart: str, expected: str
) -> None:
    neutral_payload = json.loads(_NEUTRAL.read_text(encoding="utf-8"))
    plan_payload = json.loads(_PLAN.read_text(encoding="utf-8"))
    source = neutral_payload["document"]["authored_markdown"] + "\nFirst[^a], repeated[^a].\n\n[^a]: Note body.\n"
    neutral_payload["document"]["authored_markdown"] = source
    neutral_payload["source_sha256"] = plan_payload["source_sha256"] = _sha_text(source)
    neutral, plan = tmp_path / "neutral.json", tmp_path / "plan.json"
    neutral.write_text(json.dumps(neutral_payload), encoding="utf-8")
    plan.write_text(json.dumps(plan_payload), encoding="utf-8")
    template = Document()
    template.add_paragraph("{{正文}}")
    properties = etree.SubElement(template.sections[0]._sectPr, qn("w:footnotePr"))
    etree.SubElement(properties, qn("w:numFmt"), {qn("w:val"): "decimal"})
    etree.SubElement(properties, qn("w:numRestart"), {qn("w:val"): restart})
    template_path = tmp_path / "template.docx"
    template.save(str(template_path))
    context, _workspace = _context(
        tmp_path,
        _refs(neutral, plan),
        options={"locale": "zh_CN", "heading_merge_mode": "never", "template_name": str(template_path)},
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    warnings = [item for item in result.diagnostics if item.code == "MD2DOCX-NOTE-FIELD-UPDATE-REQUIRED"]
    assert len(warnings) == (1 if restart == "eachPage" else 0)
    assert all(item.level == "warning" for item in warnings)
    with ZipFile(result.artifacts[0].staging_path) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    fields = [node for node in document.iter(qn("w:instrText")) if " NOTEREF " in (node.text or "")]
    assert len(fields) == 1
    assert fields[0].xpath("../following-sibling::w:r[2]/w:t/text()", namespaces=namespace) == [expected]
    assert fields[0].xpath("../preceding-sibling::w:r[1]/w:fldChar/@w:dirty", namespaces=namespace) == (
        ["true"] if restart == "eachPage" else []
    )
    assert len(document.findall(".//" + qn("w:footnoteReference"))) == 1
