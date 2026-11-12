"""Native caption counters reserve identities across all template stories."""

from __future__ import annotations

import re
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.opc.packuri import PackURI
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_core.docx_caption_sequences import (
    allocate_caption_sequence_scope,
    caption_sequence_identifier,
    caption_sequence_kind,
)
from docwen_core.docx_semantics import DocxSemanticRenderer
from docwen_core.models.semantic_document import SemanticCaption

pytestmark = pytest.mark.contract


def _complex_sequence(paragraph, *instruction_fragments: str) -> None:
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    paragraph.add_run()._r.append(begin)
    for fragment in instruction_fragments:
        instruction = OxmlElement("w:instrText")
        instruction.set(qn("xml:space"), "preserve")
        instruction.text = fragment
        paragraph.add_run()._r.append(instruction)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    paragraph.add_run()._r.append(separate)
    paragraph.add_run("1")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(end)


def _simple_sequence(paragraph, instruction: str) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


@pytest.mark.parametrize(
    ("kind", "prefix"),
    [("figure", "F"), ("table", "T"), ("equation", "E"), ("code_block", "C")],
)
def test_scope_binds_each_logical_kind_to_a_distinct_native_counter(kind: str, prefix: str) -> None:
    scope = allocate_caption_sequence_scope(Document(), "one request")
    identifier = caption_sequence_identifier(scope, kind)

    assert re.fullmatch(r"[0-9a-f]{24}", scope)
    assert identifier == f"DW_{prefix}_{scope}"
    assert len(identifier) <= 40
    assert caption_sequence_kind(identifier) == kind
    assert allocate_caption_sequence_scope(Document(), "one request") == scope
    assert allocate_caption_sequence_scope(Document(), "another request") != scope


@pytest.mark.parametrize("scope", ["", "a" * 23, "a" * 25, "A" * 24, "g" * 24, "a" * 24 + "\n"])
def test_scope_rejects_noncanonical_identity(scope: str) -> None:
    with pytest.raises(ValueError, match="scope"):
        caption_sequence_identifier(scope, "table")


def test_counter_rejects_unknown_logical_kind() -> None:
    with pytest.raises(ValueError, match="kind"):
        caption_sequence_identifier("a" * 24, "heading")


@pytest.mark.parametrize("identifier", ["Table", "DW_X_" + "a" * 24, "DW_T_" + "a" * 23, "DW_T_" + "A" * 24])
def test_counter_syntax_does_not_classify_conventional_or_malformed_names(identifier: str) -> None:
    assert caption_sequence_kind(identifier) is None


@pytest.mark.parametrize("kind", ["figure", "table", "equation", "code_block"])
def test_any_existing_kind_reserves_the_entire_candidate_scope(kind: str) -> None:
    document = Document()
    candidate = allocate_caption_sequence_scope(document, "template collision")
    identifier = caption_sequence_identifier(candidate, kind)
    _simple_sequence(document.add_paragraph(), f" SEQ {identifier} \\* ARABIC ")

    selected = allocate_caption_sequence_scope(document, "template collision")

    assert selected != candidate
    assert allocate_caption_sequence_scope(document, "template collision") == selected


@pytest.mark.parametrize("field_form", ["complex", "split", "simple", "lowercase", "quoted"])
def test_existing_fields_reserve_case_insensitive_identifier_across_native_forms(field_form: str) -> None:
    document = Document()
    candidate = allocate_caption_sequence_scope(document, "existing native field")
    identifier = caption_sequence_identifier(candidate, "table")
    paragraph = document.add_paragraph("Template counter ")
    if field_form == "simple":
        _simple_sequence(paragraph, f" SEQ {identifier} \\* ARABIC ")
    elif field_form == "split":
        _complex_sequence(paragraph, " S", "EQ ", identifier[:10], identifier[10:], " \\* ARABIC ")
    elif field_form == "lowercase":
        _complex_sequence(paragraph, f" seq {identifier.lower()} \\* ARABIC ")
    elif field_form == "quoted":
        _complex_sequence(paragraph, f' SEQ "{identifier}" \\* ARABIC ')
    else:
        _complex_sequence(paragraph, f" SEQ {identifier} \\* ARABIC ")
    original_xml = paragraph._p.xml

    assert allocate_caption_sequence_scope(document, "existing native field") != candidate
    assert paragraph._p.xml == original_xml


def test_nondefault_header_part_reserves_its_counter_before_and_after_reopen(tmp_path: Path) -> None:
    document = Document()
    candidate = allocate_caption_sequence_scope(document, "header story")
    identifier = caption_sequence_identifier(candidate, "figure")
    header = document.sections[0].header
    paragraph = header.paragraphs[0]
    _complex_sequence(paragraph, " SE", "Q ", identifier, " \\* ARABIC ")
    header.part._partname = PackURI("/word/stories/report-head.xml")
    selected = allocate_caption_sequence_scope(document, "header story")
    output = tmp_path / "custom-header.docx"
    document.save(str(output))

    with ZipFile(output) as package:
        assert "word/stories/report-head.xml" in package.namelist()
    reopened = Document(str(output))
    assert selected != candidate
    assert allocate_caption_sequence_scope(reopened, "header story") == selected
    assert reopened.sections[0].header.paragraphs[0]._p.xml == paragraph._p.xml


def test_reopened_generated_package_reserves_existing_request_sequence(tmp_path: Path) -> None:
    document = Document()
    candidate = allocate_caption_sequence_scope(document, "generated package")
    DocxSemanticRenderer(document, sequence_scope=candidate).render_caption(
        SemanticCaption("table", "tbl-existing", "1", "Table", "Existing generated caption")
    )
    output = tmp_path / "generated.docx"
    document.save(str(output))
    reopened = Document(str(output))

    next_scope = allocate_caption_sequence_scope(reopened, "generated package")

    assert next_scope != candidate
    assert next_scope == allocate_caption_sequence_scope(document, "generated package")
    assert caption_sequence_identifier(candidate, "table") in reopened.element.xml
