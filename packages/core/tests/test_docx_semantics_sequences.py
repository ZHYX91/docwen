"""Neutral captions use native fields without advancing template sequences."""

from __future__ import annotations

import re

import pytest
from docx import Document
from docx.oxml.ns import qn

from docwen_core.docx_bookmarks import build_docx_bookmark_inventory
from docwen_core.docx_parsing.document_semantics import encode_target_bookmark, extract_neutral_semantic_caption
from docwen_core.docx_semantics import DocxSemanticImporter, DocxSemanticRenderer, append_complex_field
from docwen_core.models.semantic_document import (
    SemanticCaption,
    SemanticCaptionKind,
    SemanticReference,
    SemanticTable,
    SemanticTableCell,
)

pytestmark = pytest.mark.contract


def _table(target: str, number: str) -> SemanticTable:
    return SemanticTable(
        row_count=1,
        column_count=1,
        cells=(SemanticTableCell(0, 0, target),),
        caption=SemanticCaption("table", target, number, "表", target),
    )


def _instructions(document) -> list[str]:
    return [(item.text or "").strip() for item in document.element.iter(qn("w:instrText"))]


def test_neutral_caption_sequence_isolated_from_template_before_and_after_save(tmp_path) -> None:
    document = Document()
    before = document.add_paragraph("Template before ")
    append_complex_field(before, instruction=" SEQ Table \\* ARABIC ", cached_result="1")
    before_xml = before._p.xml
    renderer = DocxSemanticRenderer(document)
    first = _table("tbl-first", "1")
    second = _table("tbl-second", "2")
    renderer.render_table(first)
    renderer.render_table(second)
    reference = document.add_paragraph("See ")
    renderer.render_reference(reference, SemanticReference("tbl-second", "2"))
    after = document.add_paragraph("Template after ")
    append_complex_field(after, instruction=" SEQ Table \\* ARABIC ", cached_result="2")
    after_xml = after._p.xml
    output = tmp_path / "independent-captions.docx"
    document.save(output)
    reopened = Document(output)

    instructions = _instructions(reopened)
    generated_identifier = instructions[1].split()[1]
    assert re.fullmatch(r"DW_T_[0-9a-f]{24}", generated_identifier)
    assert instructions == [
        "SEQ Table \\* ARABIC",
        f"SEQ {generated_identifier} \\* ARABIC",
        f"SEQ {generated_identifier} \\* ARABIC",
        f"REF {encode_target_bookmark('tbl-second')} \\h \\* MERGEFORMAT",
        "SEQ Table \\* ARABIC",
    ]
    assert reopened.paragraphs[0]._p.xml == before_xml
    assert reopened.paragraphs[-1]._p.xml == after_xml
    imported = DocxSemanticImporter().import_document(reopened)
    assert imported.diagnostics == ()
    assert [block for block in imported.document.blocks if isinstance(block, SemanticTable)] == [first, second]


def test_independent_renderers_allocate_sequences_when_first_used() -> None:
    document = Document()
    first_renderer = DocxSemanticRenderer(document)
    second_renderer = DocxSemanticRenderer(document)
    first_renderer.render_table(_table("tbl-first", "1"))
    second_renderer.render_table(_table("tbl-second", "1"))
    first_renderer.render_table(_table("tbl-third", "2"))

    identifiers = [instruction.split()[1] for instruction in _instructions(document)]
    assert all(re.fullmatch(r"DW_T_[0-9a-f]{24}", identifier) for identifier in identifiers)
    assert identifiers[0] == identifiers[2]
    assert identifiers[0] != identifiers[1]


def test_explicit_scope_shares_request_sequence_across_renderers() -> None:
    document = Document()
    scope = "0123456789abcdef01234567"
    DocxSemanticRenderer(document, sequence_scope=scope).render_table(_table("tbl-first", "1"))
    DocxSemanticRenderer(document, sequence_scope=scope).render_table(_table("tbl-second", "2"))

    assert _instructions(document) == [f"SEQ DW_T_{scope} \\* ARABIC"] * 2


@pytest.mark.parametrize("scope", ["", "a" * 23, "g" * 24, "a" * 25])
def test_explicit_scope_rejects_invalid_physical_identity(scope: str) -> None:
    document = Document()
    original_xml = document.element.xml

    with pytest.raises(ValueError):
        DocxSemanticRenderer(document, sequence_scope=scope)

    assert document.element.xml == original_xml


@pytest.mark.parametrize(("kind", "prefix"), [("figure", "F"), ("table", "T"), ("equation", "E"), ("listing", "C")])
def test_physical_sequence_kind_round_trips_with_caption_target(kind: SemanticCaptionKind, prefix: str) -> None:
    document = Document()
    caption = SemanticCaption(kind, "target", "4", "本地化标签", "Caption body")
    paragraph = DocxSemanticRenderer(document).render_caption(caption)

    identifier = _instructions(document)[0].split()[1]
    assert re.fullmatch(rf"DW_{prefix}_[0-9a-f]{{24}}", identifier)
    extracted = extract_neutral_semantic_caption(
        paragraph._p,
        bookmark_inventory=build_docx_bookmark_inventory(document),
    )
    assert extracted is not None
    assert (extracted.kind, extracted.target_id, extracted.label, extracted.cached_number, extracted.content) == (
        kind,
        "target",
        "本地化标签",
        "4",
        "Caption body",
    )


def test_physical_sequence_identifier_alone_does_not_claim_caption_ownership() -> None:
    document = Document()
    paragraph = document.add_paragraph("Table ")
    append_complex_field(
        paragraph,
        instruction=" SEQ DW_T_0123456789abcdef01234567 \\* ARABIC ",
        cached_result="1",
    )
    paragraph.add_run(": User-owned caption")

    assert (
        extract_neutral_semantic_caption(
            paragraph._p,
            bookmark_inventory=build_docx_bookmark_inventory(document),
            proven_object_kind="table",
        )
        is None
    )
