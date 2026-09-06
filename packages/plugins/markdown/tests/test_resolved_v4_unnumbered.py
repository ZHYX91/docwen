"""Resolved disabled targets retain visible title/Alias references without fields."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx.oxml.ns import qn
from lxml import etree

from docwen_core.models.resolved_numbering import canonicalize_numbering_plan
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from .test_resolved_v4_to_docx import _NEUTRAL, _PLAN, _context, _refs

pytestmark = pytest.mark.contract


def test_unnumbered_resolved_targets_render_titles_and_recover_references(tmp_path: Path) -> None:
    from docx import Document

    from docwen_core.docx_resolved_numbering_recovery import ResolvedNumberingV4Recovery

    neutral = json.loads(_NEUTRAL.read_text(encoding="utf-8"))
    plan = json.loads(_PLAN.read_text(encoding="utf-8"))
    plan["plan"]["heading_definitions"] = []
    plan["plan"]["heading_instances"] = []
    for target in plan["plan"]["targets"]:
        target.update(enabled=False, derived_number=None, materialization=None)
    for reference in neutral["document"]["references"]:
        reference["cached_number"] = ""
    digest = hashlib.sha256(canonicalize_numbering_plan(plan["plan"])).hexdigest()
    neutral["plan_sha256"] = plan["plan_sha256"] = digest
    neutral_path = tmp_path / "neutral.json"
    plan_path = tmp_path / "plan.json"
    neutral_path.write_text(json.dumps(neutral), encoding="utf-8")
    plan_path.write_text(json.dumps(plan), encoding="utf-8")
    context, _workspace = _context(tmp_path, _refs(neutral_path, plan_path))
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    output = Path(result.artifacts[0].staging_path)
    with ZipFile(output) as package:
        document = etree.fromstring(package.read("word/document.xml"))
        instructions = [node.text or "" for node in document.iter(qn("w:instrText"))]
        assert not any(" REF " in value or "SEQ " in value for value in instructions)
        visible = "".join(node.text or "" for node in document.iter(qn("w:t")))
        assert "Stable: Architecture and System overview." in visible
    reopened = Document(str(output))
    recovery = ResolvedNumberingV4Recovery.load_if_present(output, reopened)
    assert recovery is not None
    recovered = "\n".join(recovery.render_paragraph_text(node) or "" for node in document.iter(qn("w:p")))
    assert "@[[#^h-7f3a]]" in recovered
    assert "@[[#^system-overview|System overview]]" in recovered
