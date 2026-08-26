"""Fail-closed unit gates for the packaged Machine exact-two document-semantics probe."""

from __future__ import annotations

import io
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from scripts.release import verify_packaged_cli as packaged_cli
from scripts.release.verify_packaged_cli import (
    MACHINE_EXACT_TWO_IMAGE_BYTES,
    MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT,
    MACHINE_EXACT_TWO_NOTE_ROUNDTRIP_TOKENS,
    MACHINE_RESOLVED_DOCUMENT_LIMITATIONS,
    verify_machine_document_semantics_docx,
    verify_machine_document_semantics_markdown,
    verify_machine_note_domains_markdown,
)

pytestmark = pytest.mark.unit

_AUTH_MARKDOWN = MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT["document"]["authored_markdown"]

_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
   <w:bookmarkStart w:id="1" w:name="DW_T_00000000000000000000000000000000000"/><w:r><w:t>Architecture</w:t></w:r>
   <w:bookmarkEnd w:id="1"/></w:p>
  <w:p><w:pPr><w:pStyle w:val="DocWenFigureCaption"/></w:pPr>
   <w:bookmarkStart w:id="2" w:name="DW_T_00000000000000000000000000000000001"/><w:r><w:t>Figure</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure \\* ARABIC </w:instrText></w:r>
   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> System overview</w:t></w:r>
   <w:bookmarkEnd w:id="2"/></w:p>
  <w:p><w:r><w:drawing><wp:docPr id="1" name="Picture 1"/><wp:blip r:embed="rIdImage"/></w:drawing></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="DocWenTableCaption"/></w:pPr>
   <w:bookmarkStart w:id="3" w:name="DW_T_00000000000000000000000000000000002"/><w:r><w:t>Table</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Table \\* ARABIC </w:instrText></w:r>
   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> Results</w:t></w:r>
   <w:bookmarkEnd w:id="3"/></w:p>
  <w:tbl><w:tr><w:tc><w:tcPr><w:tblHeader/></w:tcPr><w:tc><w:p><w:r><w:t>Metric</w:t></w:r></w:p></w:tc>
   <w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tc></w:tr></w:tbl>
  <w:p><w:pPr><w:pStyle w:val="DocWenEquationCaption"/></w:pPr>
   <w:bookmarkStart w:id="4" w:name="DW_T_00000000000000000000000000000000003"/><w:r><w:t>Equation</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Equation \\* ARABIC </w:instrText></w:r>
   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="4"/></w:p>
  <w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
   <m:r><m:t>E=mc^2</m:t></m:r></m:oMath></w:p>
  <w:p><w:pPr><w:pStyle w:val="DocWenCodeBlockCaption"/></w:pPr>
   <w:bookmarkStart w:id="5" w:name="DW_T_00000000000000000000000000000000004"/><w:r><w:t>Code</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Code \\* ARABIC </w:instrText></w:r>
   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> Entry point</w:t></w:r>
   <w:bookmarkEnd w:id="5"/></w:p>
  <w:p><w:r><w:t>fn main() {}</w:t></w:r></w:p>
  <w:p><w:r><w:t>Stable: </w:t></w:r>
   <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF DW_T_00000000000000000000000000000000000 \\n \\h </w:instrText></w:r>
   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="end"/></w:r>
   <w:r><w:t> and </w:t></w:r>
   <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF DW_T_00000000000000000000000000000000001 \\h </w:instrText></w:r>
   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">|System overview</w:t></w:r>
   <w:r><w:t>.</w:t></w:r></w:p>
  <w:p><w:r><w:t>Ordinary: [[#^system-overview]] and ![[Guide#^h-7f3a]].</w:t></w:r></w:p>
  <w:p><w:r><w:t>Citation: </w:t></w:r>
   <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> CITATION cite-one </w:instrText></w:r>
   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>One (2026)</w:t></w:r>
   <w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>.</w:t></w:r></w:p>
 </w:body>
</w:document>
"""

_RELATIONSHIPS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rIdImage"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  Target="media/image1.png"/>
</Relationships>
"""


def _write_probe(
    path: Path,
    *,
    document_xml: str = _DOCUMENT_XML,
    relationships_xml: str = _RELATIONSHIPS_XML,
    include_media: bool = True,
) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/_rels/document.xml.rels", relationships_xml)
        if include_media:
            archive.writestr("word/media/image1.png", MACHINE_EXACT_TWO_IMAGE_BYTES)


def test_packaged_machine_semantic_oracles_accept_the_complete_probe(tmp_path: Path) -> None:
    docx_path = tmp_path / "complete.docx"
    _write_probe(docx_path)

    verify_machine_document_semantics_docx(docx_path)
    verify_machine_document_semantics_markdown(_AUTH_MARKDOWN)
    assert [item["code"] for item in MACHINE_RESOLVED_DOCUMENT_LIMITATIONS] == [
        "resolved_document.provider_owned_semantics",
    ]


def test_packaged_machine_note_oracle_accepts_the_frozen_domain_projection() -> None:
    verify_machine_note_domains_markdown("\n".join(MACHINE_EXACT_TWO_NOTE_ROUNDTRIP_TOKENS))


@pytest.mark.parametrize(
    "definition",
    MACHINE_EXACT_TWO_NOTE_ROUNDTRIP_TOKENS[1:],
)
def test_packaged_machine_note_oracle_rejects_extra_definition_separator(definition: str) -> None:
    markdown = "\n".join(MACHINE_EXACT_TWO_NOTE_ROUNDTRIP_TOKENS).replace(
        definition,
        definition.replace(": ", ":  "),
    )

    with pytest.raises(RuntimeError, match="note_domains_markdown_missing"):
        verify_machine_note_domains_markdown(markdown)


@pytest.mark.parametrize("token", MACHINE_EXACT_TWO_NOTE_ROUNDTRIP_TOKENS)
def test_packaged_machine_note_oracle_rejects_each_domain_regression(token: str) -> None:
    markdown = "\n".join(MACHINE_EXACT_TWO_NOTE_ROUNDTRIP_TOKENS).replace(token, "missing", 1)

    with pytest.raises(RuntimeError, match="note_domains_markdown_missing"):
        verify_machine_note_domains_markdown(markdown)


def test_packaged_machine_docx_oracle_does_not_reopen_the_path(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    docx_path = tmp_path / "path-read-once.docx"
    _write_probe(docx_path)
    original = packaged_cli.zipfile.ZipFile

    def require_in_memory_archive(value: object, *args: object, **kwargs: object) -> ZipFile:
        assert isinstance(value, io.BytesIO)
        return original(value, *args, **kwargs)

    monkeypatch.setattr(packaged_cli.zipfile, "ZipFile", require_in_memory_archive)
    packaged_cli.verify_machine_document_semantics_docx(docx_path)


@pytest.mark.parametrize(
    ("old", "new", "error"),
    [
        (
            'w:name="DW_T_00000000000000000000000000000000000"',
            'w:name="ordinary"',
            "target_bookmarks_invalid",
        ),
        (
            "REF DW_T_00000000000000000000000000000000000",
            "REF DW_T_99999999999999999999999999999999999",
            "reference_bookmark_missing",
        ),
        ("SEQ Figure", "SEQ Picture", "seq_invalid"),
        ("SEQ Table", "SEQ Grid", "seq_invalid"),
        (
            '<w:instrText> REF DW_T_00000000000000000000000000000000001 \\h </w:instrText></w:r>\n   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>',
            '<w:instrText> REF DW_T_00000000000000000000000000000000001 \\h </w:instrText></w:r>\n   <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r>',
            "reference_cached_mismatch",
        ),
        ("Architecture", "heading missing", "visible_text_missing"),
        ("Entry point", "caption missing", "visible_text_missing"),
    ],
)
def test_packaged_machine_docx_oracle_rejects_each_semantic_regression(
    tmp_path: Path,
    old: str,
    new: str,
    error: str,
) -> None:
    docx_path = tmp_path / "mutated.docx"
    _write_probe(docx_path, document_xml=_DOCUMENT_XML.replace(old, new, 1))

    with pytest.raises(RuntimeError, match=error):
        verify_machine_document_semantics_docx(docx_path)


def test_packaged_machine_docx_oracle_rejects_a_missing_media(tmp_path: Path) -> None:
    docx_path = tmp_path / "missing-media.docx"
    _write_probe(docx_path, include_media=False)

    with pytest.raises(RuntimeError, match="media_missing"):
        verify_machine_document_semantics_docx(docx_path)


def test_packaged_machine_docx_oracle_rejects_a_disabled_occurrence(tmp_path: Path) -> None:
    docx_path = tmp_path / "disabled-occurrence.docx"
    disabled = _DOCUMENT_XML.replace(
        "<w:body>",
        '<w:body><w:sdt><w:sdtPr><w:tag w:val="docwen-numbering-occurrence-v1:00000000000000000000000000000000"/></w:sdtPr></w:sdt>',
        1,
    )
    _write_probe(docx_path, document_xml=disabled)

    with pytest.raises(RuntimeError, match="unexpected_disabled_occurrence"):
        verify_machine_document_semantics_docx(docx_path)


@pytest.mark.parametrize(
    "token",
    [
        "# Architecture ^h-7f3a",
        "Figure: System overview ^system-overview",
        "Table: Results ^results-main",
        "Equation: ^energy-main",
        "Code: Entry point ^entry-main",
        "Stable: @[[#^h-7f3a]] and @[[#^system-overview|System overview]].",
        "Ordinary: [[#^system-overview]] and ![[Guide#^h-7f3a]].",
        "Citation: @cite-one.",
    ],
)
def test_packaged_machine_markdown_oracle_rejects_each_semantic_regression(token: str) -> None:
    with pytest.raises(RuntimeError, match="markdown_missing"):
        verify_machine_document_semantics_markdown(_AUTH_MARKDOWN.replace(token, "missing", 1))


def test_packaged_machine_markdown_oracle_rejects_a_missing_image() -> None:
    markdown = _AUTH_MARKDOWN.replace("![[system.png]]", "image missing")

    with pytest.raises(RuntimeError, match="markdown_image_missing"):
        verify_machine_document_semantics_markdown(markdown)
