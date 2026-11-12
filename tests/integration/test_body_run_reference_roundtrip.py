"""Direct body formatting and authenticated references retain separate owners."""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.text.run import CT_R
from docx.shared import Pt, RGBColor
from docx.text.run import Run
from lxml import etree

from docwen_core.docx_bookmarks import build_docx_bookmark_inventory, prove_bookmark_name
from docwen_core.docx_citation_ooxml import (
    build_resolved_citation_projection,
    citation_instruction,
    read_proven_resolved_citations,
)
from docwen_core.docx_semantics_v3 import DocxSemanticsV3Recovery
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    NUMBERING_EXPORT_PLAN_SCHEMA,
    NUMBERING_EXPORT_PLAN_SCHEMA_ID,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
    RESOLVED_DOCUMENT_SCHEMA,
    RESOLVED_DOCUMENT_SCHEMA_ID,
    NumberingExportPlanEnvelope,
    ResolvedCitation,
    ResolvedCitationItem,
    ResolvedDocument,
    ResolvedDocumentEnvelope,
    ResolvedNumberingPlan,
    canonicalize_numbering_plan,
)
from docwen_runtime.templates import TemplateManager
from tests.integration._round_trip_helper import _primary_path, docx_to_md, md_to_docx

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def test_body_anchor_formatting_preserves_reference_fields_aliases_and_soft_tokens(
    round_trip_runtime: Any, tmp_path: Path
) -> None:
    template = Document()
    existing = template.add_paragraph("Template prefix stays unchanged.")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), " DATE ")
    field_run = OxmlElement("w:r")
    field_text = OxmlElement("w:t")
    field_text.text = "2026-11-12"
    field_run.append(field_text)
    field.append(field_run)
    existing._p.append(field)
    original_existing = etree.tostring(existing._p, method="c14n", exclusive=True)
    body_style = template.styles.add_style("Contrasting Body", WD_STYLE_TYPE.PARAGRAPH)
    body_style.font.name = "Cambria"
    body_style.font.size = Pt(22)
    body_style.font.color.rgb = RGBColor.from_string("AA6622")
    anchor = template.add_paragraph("{{body}}", style=body_style)
    anchor.runs[0].font.name = "Calibri"
    anchor.runs[0].font.size = Pt(11)
    anchor.runs[0].font.color.rgb = RGBColor.from_string("224466")
    template_path = tmp_path / "template.docx"
    template.save(str(template_path))
    template_hash = hashlib.sha256(template_path.read_bytes()).hexdigest()
    imported = TemplateManager.default().import_template(template_path)
    source = tmp_path / "source.md"
    source.write_text(
        "# Unnumbered\n\n"
        "Table: Results ^result\n\n| Name | Value |\n| --- | --- |\n| Synthetic | 23 |\n\n"
        "Ordinary **strong** and *italic* [link](https://example.org/qa). "
        "REF **@[[#^result|result alias]]**, soft *@[[#Unnumbered|soft alias]]*.\n",
        encoding="utf-8",
    )
    output = md_to_docx(
        round_trip_runtime,
        source,
        tmp_path / "docx-output",
        options={"template_name": imported.id, "heading_merge_mode": "never"},
    )
    document = Document(str(output))
    recovery = DocxSemanticsV3Recovery.load(output, document)
    assert recovery.caption_signatures == (("table", "result", "Results", "1"),)
    assert len(recovery.reference_occurrence_identities) == 1
    assert len(recovery.soft_reference_identities) == 1
    [target] = recovery.target_identities
    bookmark = prove_bookmark_name(build_docx_bookmark_inventory(document), target.bookmark_name)
    assert bookmark.valid

    paragraph = next(item for item in document.paragraphs if item.text.startswith("Ordinary "))
    runs = [Run(element, paragraph) for element in paragraph._p.iter(qn("w:r"))]
    for text in ("Ordinary ", "strong", "italic", "link"):
        run = next(item for item in runs if item.text == text)
        assert run.font.name == "Calibri" and run.font.size == Pt(11)
        if text != "link":
            assert run.font.color.rgb == RGBColor.from_string("224466")
    assert next(item for item in runs if item.text == "strong").bold is True
    assert next(item for item in runs if item.text == "italic").italic is True
    assert next(item for item in runs if item.text == "link").style.style_id == "Hyperlink"
    assert paragraph.style is not None and paragraph.style.style_id == "ContrastingBody"
    assert paragraph.style.font.name == "Cambria" and paragraph.style.font.size == Pt(22)
    assert paragraph.style.font.color.rgb == RGBColor.from_string("AA6622")
    controls = list(paragraph._p.iter(qn("w:sdt")))
    assert len(controls) == 2
    for index, control in enumerate(controls):
        for element in control.iter(qn("w:r")):
            run = Run(element, paragraph)
            if element.find(qn("w:t")) is None:
                assert element.find(qn("w:rPr")) is None
                continue
            assert run.font.name == "Calibri" and run.font.size == Pt(11)
            assert run.font.color.rgb == RGBColor.from_string("224466")
            assert (run.bold if index == 0 else run.italic) is True
    [instruction] = list(controls[0].iter(qn("w:instrText")))
    assert instruction.text == f" REF {target.bookmark_name} \\h \\* MERGEFORMAT "
    assert [node.text for node in controls[0].iter(qn("w:t"))] == ["1", " result alias"]
    assert document.paragraphs[0].text == "Template prefix stays unchanged."
    assert etree.tostring(document.paragraphs[0]._p, method="c14n", exclusive=True) == original_existing
    assert document.paragraphs[0].runs[0]._r.rPr is None
    assert hashlib.sha256(template_path.read_bytes()).hexdigest() == template_hash
    assert hashlib.sha256(imported.path.read_bytes()).hexdigest() == template_hash

    returned = docx_to_md(round_trip_runtime, output, tmp_path / "markdown-output", preserve_numbering=False)
    assert "Table: Results ^result" in returned
    assert "**strong**" in returned and "*italic*" in returned
    assert "@[[#^result|result alias]]" in returned
    assert "@[[#Unnumbered|soft alias]]" in returned


def test_body_anchor_formatting_preserves_locked_resolved_citation_after_save(
    round_trip_runtime: Any, tmp_path: Path
) -> None:
    token = "@cite-font"
    authored = f"Evidence **{token}**.\n"
    source_digest = hashlib.sha256(authored.encode()).hexdigest()
    citation = ResolvedCitation(
        source_start=authored.index(token),
        source_end=authored.index(token) + len(token),
        source_slice_sha256=hashlib.sha256(token.encode()).hexdigest(),
        authored_token=token,
        form="narrative",
        cluster_id="cluster-format",
        items=(
            ResolvedCitationItem(
                "cite-font",
                "reference-record:98",
                hashlib.sha256(b"synthetic-record").hexdigest(),
                "Synthetic Author (2026)",
            ),
        ),
        cached_result="Synthetic Author (2026)",
    )
    plan = ResolvedNumberingPlan((), (), ())
    plan_digest = hashlib.sha256(
        canonicalize_numbering_plan({"heading_definitions": [], "heading_instances": [], "targets": []})
    ).hexdigest()
    resolved = ResolvedDocument(authored, (), (), (), (citation,), ())
    resources = (
        (
            "resolved.json",
            RESOLVED_DOCUMENT_SCHEMA_ID,
            RESOLVED_DOCUMENT_SCHEMA,
            ResolvedDocumentEnvelope("citation-format", source_digest, plan_digest, resolved),
        ),
        (
            "plan.json",
            NUMBERING_EXPORT_PLAN_SCHEMA_ID,
            NUMBERING_EXPORT_PLAN_SCHEMA,
            NumberingExportPlanEnvelope("citation-format", source_digest, plan_digest, plan),
        ),
    )
    for name, schema_id, schema, envelope in resources:
        (tmp_path / name).write_text(
            json.dumps({"$schema": schema_id, "schema": schema, **asdict(envelope)}), encoding="utf-8"
        )
    refs = [
        FileRef(
            path=str(tmp_path / "resolved.json"),
            format="markdown",
            category="document",
            input_kind="document",
            input_role="neutral_document",
            logical_path="citation.md",
            media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
        ),
        FileRef(
            path=str(tmp_path / "plan.json"),
            format="json",
            category="resource",
            input_kind="resource",
            input_role="numbering_export_plan",
            logical_path="plan.json",
            media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
        ),
    ]
    template = Document()
    body_style = template.styles.add_style("Citation Body", WD_STYLE_TYPE.PARAGRAPH)
    body_style.font.name = "Cambria"
    body_style.font.size = Pt(22)
    body_style.font.color.rgb = RGBColor.from_string("AA6622")
    anchor = template.add_paragraph("{{body}}", style=body_style)
    anchor.runs[0].font.name = "Calibri"
    anchor.runs[0].font.size = Pt(11)
    anchor.runs[0].font.color.rgb = RGBColor.from_string("224466")
    template_path = tmp_path / "citation-template.docx"
    template.save(str(template_path))
    imported = TemplateManager.default().import_template(template_path)
    result = round_trip_runtime.execute(
        ConversionRequest(
            request_id="body-citation-format",
            input_refs=refs,
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "citation-output")),
            options={"template_name": imported.id, "heading_merge_mode": "never"},
        )
    )
    output = _primary_path(result)
    document = Document(str(output))
    paragraph = next(item for item in document.paragraphs if item.text.startswith("Evidence "))
    assert paragraph.style is not None and paragraph.style.font.name == "Cambria"
    assert paragraph.style.font.size == Pt(22)
    assert paragraph.style.font.color.rgb == RGBColor.from_string("AA6622")
    [control] = list(paragraph._p.iter(qn("w:sdt")))
    [cached] = list(control.iter(qn("w:t")))
    assert cached.text == citation.cached_result
    for element in control.iter(qn("w:r")):
        assert isinstance(element, CT_R)
        if element.find(qn("w:t")) is None:
            assert element.find(qn("w:rPr")) is None
            continue
        run = Run(element, paragraph)
        assert run.font.name == "Calibri" and run.font.size == Pt(11)
        assert run.font.color.rgb == RGBColor.from_string("224466")
        assert run.bold is True
    markers = list(control.iter(qn("w:fldChar")))
    assert [marker.get(qn("w:fldCharType")) for marker in markers] == ["begin", "separate", "end"]
    assert markers[0].get(qn("w:fldLock")) == "true"
    assert all(marker.get(qn("w:dirty")) is None for marker in markers)
    projection = build_resolved_citation_projection(source_digest, (citation,))
    assert projection is not None
    [occurrence] = projection.occurrence_map.occurrences
    [instruction] = list(control.iter(qn("w:instrText")))
    assert instruction.text == citation_instruction(occurrence.item_refs)
    assert read_proven_resolved_citations(output) == (citation,)
    saved_again = tmp_path / "citation-saved-again.docx"
    document.save(str(saved_again))
    assert read_proven_resolved_citations(saved_again) == (citation,)
    returned = docx_to_md(round_trip_runtime, saved_again, tmp_path / "citation-markdown")
    assert token in returned
    assert citation.cached_result not in returned
