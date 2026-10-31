"""Callout comments and note fallbacks through the resolved-v4 file route."""

from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx.oxml.ns import qn
from lxml import etree

from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from ._resolved_v4_support import _NEUTRAL, _PLAN, _context, _refs, _sha_text

pytestmark = pytest.mark.contract


def test_exact_two_callout_note_fallback_reaches_result_and_saved_note(tmp_path: Path) -> None:
    neutral_payload = json.loads(_NEUTRAL.read_text(encoding="utf-8"))
    plan_payload = json.loads(_PLAN.read_text(encoding="utf-8"))
    source = neutral_payload["document"]["authored_markdown"] + (
        "\n> [!tip] Main title\n> Body\n\nText[^1]\n\n[^1]: > [!warning] Check\n    > Save first\n"
    )
    neutral_payload["source_sha256"] = plan_payload["source_sha256"] = _sha_text(source)
    neutral_payload["document"]["authored_markdown"] = source
    neutral, plan = tmp_path / "neutral.json", tmp_path / "plan.json"
    neutral.write_text(json.dumps(neutral_payload), encoding="utf-8")
    plan.write_text(json.dumps(plan_payload), encoding="utf-8")
    context, _workspace = _context(tmp_path, _refs(neutral, plan))
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    warnings = [item for item in result.diagnostics if item.code == "CALLOUT-NOTE-TEXT"]
    assert len(warnings) == 1 and "(1)" in warnings[0].message
    with ZipFile(result.artifacts[0].staging_path) as package:
        notes = etree.fromstring(package.read("word/footnotes.xml"))
        visible = "".join(item.text or "" for item in notes.iter(qn("w:t")))
        assert "Warning: Check" in visible and "Save first" in visible
        assert "[!warning]" not in visible
        comments = etree.fromstring(package.read("word/comments.xml"))
        assert "Markdown callout type: tip" in "".join(item.text or "" for item in comments.iter(qn("w:t")))
