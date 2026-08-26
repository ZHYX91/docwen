"""Focused tests split from test_docx_citations.py."""

from __future__ import annotations

from ._docx_citations_support import (
    BIBLIOGRAPHY_BOOKMARK_NAME,
    CT,
    RT,
    Any,
    Document,
    DocxSemanticImporter,
    DocxSemanticRenderer,
    Inches,
    PackURI,
    Part,
    Path,
    Pt,
    SemanticBibliographyEntry,
    SemanticBibliographyFragment,
    SemanticBibliographyRun,
    SemanticCaption,
    SemanticCitationCluster,
    SemanticDocument,
    SemanticDocumentValidationError,
    SemanticParagraph,
    SemanticTable,
    SemanticTableCell,
    SemanticText,
    _assert_citation_fell_back_to_text,
    _bookmark_range,
    _citation,
    _citation_field_char,
    _remove_element,
    _render_bibliography,
    _render_citation,
    append_bookmark_start,
    append_zero_width_bookmark,
    deepcopy,
    encode_bibliography_entry_bookmark,
    encode_citation_bookmark,
    etree,
    pytest,
    qn,
)

pytestmark = pytest.mark.contract


def test_core_docx_semantic_sources_remain_provider_neutral() -> None:
    core_root = Path(__file__).resolve().parents[1] / "src" / "docwen_core"
    source_paths = (
        core_root / "docx_semantics.py",
        core_root / "docx_parsing" / "document_semantics.py",
        core_root / "models" / "semantic_document.py",
    )
    forbidden = ("wen" + "leaf", "p" + "kwf", "c" + "sl", "@[" + "[")

    violations = {
        source_path.relative_to(core_root).as_posix(): marker
        for source_path in source_paths
        for marker in forbidden
        if marker in source_path.read_text(encoding="utf-8").casefold()
    }

    assert violations == {}


def test_multiple_citation_clusters_round_trip_with_ordered_items_and_locked_clean_fields() -> None:
    first = _citation("cluster-one", "smith2025")
    second = _citation(
        "cluster-two",
        "wang2024",
        "smith2025",
        cached_result="[2, 1]",
    )
    semantic = SemanticDocument(
        blocks=(
            SemanticParagraph(
                (
                    SemanticText("Sources "),
                    first,
                    SemanticText(" and "),
                    second,
                    SemanticText("."),
                )
            ),
        )
    )
    document = Document()
    DocxSemanticRenderer(document).render_blocks(semantic.blocks)

    instructions = [item.text for item in document.element.iter(qn("w:instrText"))]
    begins = [item for item in document.element.iter(qn("w:fldChar")) if item.get(qn("w:fldCharType")) == "begin"]
    marker_names = {item.get(qn("w:name")) for item in document.element.iter(qn("w:bookmarkStart"))}
    imported = DocxSemanticImporter().import_document(document)

    assert instructions == [
        " CITATION smith2025 ",
        r" CITATION wang2024 \m smith2025 ",
    ]
    assert all(item.get(qn("w:fldLock")) == "true" for item in begins)
    assert all(item.get(qn("w:dirty")) is None for item in begins)
    assert marker_names == {
        encode_citation_bookmark("cluster-one"),
        encode_citation_bookmark("cluster-two"),
    }
    assert imported.diagnostics == ()
    assert imported.document == semantic


def test_citation_renderer_rejects_reserved_marker_collision_before_field_write() -> None:
    document = Document()
    existing = document.add_paragraph("existing")
    marker_name = encode_citation_bookmark("cluster-one")
    append_zero_width_bookmark(existing, marker_name, "50")
    paragraph = document.add_paragraph()

    with pytest.raises(SemanticDocumentValidationError) as error:
        DocxSemanticRenderer(document).render_citation(paragraph, _citation())

    assert [item.code for item in error.value.diagnostics] == ["semantic.docx.bookmark.name_conflict"]
    assert "CITATION" not in paragraph._p.xml


def test_citation_renderer_rejects_a_paragraph_from_another_document() -> None:
    document = Document()
    foreign_document = Document()
    foreign_paragraph = foreign_document.add_paragraph("foreign")
    document_before = document.element.xml
    paragraph_before = foreign_paragraph._p.xml

    with pytest.raises(ValueError, match="must belong to the renderer document part"):
        DocxSemanticRenderer(document).render_citation(foreign_paragraph, _citation())

    assert document.element.xml == document_before
    assert foreign_paragraph._p.xml == paragraph_before


@pytest.mark.parametrize("field_type", ["begin", "separate", "end"])
def test_broken_citation_field_fails_closed_and_preserves_cached_text(field_type: str) -> None:
    document, _ = _render_citation()
    field_char = _citation_field_char(document, field_type)
    _remove_element(field_char)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.field_invalid"]
    _assert_citation_fell_back_to_text(document)


def test_unlocked_citation_fails_closed_and_preserves_cached_text() -> None:
    document, _ = _render_citation()
    begin = _citation_field_char(document, "begin")
    begin.attrib.pop(qn("w:fldLock"))

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.field_unlocked"]
    _assert_citation_fell_back_to_text(document)


@pytest.mark.parametrize("field_type", ["begin", "separate", "end"])
def test_dirty_citation_fails_closed_and_preserves_cached_text(field_type: str) -> None:
    document, _ = _render_citation()
    field_char = _citation_field_char(document, field_type)
    field_char.set(qn("w:dirty"), "true")

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.field_dirty"]
    _assert_citation_fell_back_to_text(document)


def test_invalid_citation_instruction_fails_closed_and_preserves_cached_text() -> None:
    document, _ = _render_citation()
    instruction = next(document.element.iter(qn("w:instrText")))
    instruction.text = r" CITATION smith2025 \x wang2024 "

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.instruction_invalid"]
    _assert_citation_fell_back_to_text(document)


def test_empty_citation_cached_result_fails_closed() -> None:
    document, _ = _render_citation()
    for text in document.element.iter(qn("w:t")):
        text.text = ""

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.cached_result_empty"]
    assert all(
        not isinstance(inline, SemanticCitationCluster)
        for block in imported.document.blocks
        if isinstance(block, SemanticParagraph)
        for inline in block.inlines
    )


@pytest.mark.parametrize("remove_field_end", [False, True])
def test_complete_or_incomplete_orphan_citation_preserves_the_whole_paragraph(
    remove_field_end: bool,
) -> None:
    citation = _citation()
    semantic = SemanticParagraph(
        (
            SemanticText("Before "),
            citation,
            SemanticText(" after."),
        )
    )
    document = Document()
    DocxSemanticRenderer(document).render_blocks((semantic,))
    marker_name = encode_citation_bookmark("cluster-one")
    start, end = _bookmark_range(document, marker_name)
    _remove_element(start)
    _remove_element(end)
    if remove_field_end:
        field_end = _citation_field_char(document, "end")
        _remove_element(field_end)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.marker_missing"]
    assert imported.document.blocks == (SemanticParagraph((SemanticText("Before [1] after."),)),)


def test_valid_cluster_mixed_with_an_orphan_downgrades_the_whole_paragraph() -> None:
    first = _citation("cluster-one", "smith2025", cached_result="[1]")
    second = _citation("cluster-two", "wang2024", cached_result="[2]")
    document = Document()
    DocxSemanticRenderer(document).render_blocks(
        (
            SemanticParagraph(
                (
                    SemanticText("First "),
                    first,
                    SemanticText(" then "),
                    second,
                    SemanticText("."),
                )
            ),
        )
    )
    orphan_start, orphan_end = _bookmark_range(
        document,
        encode_citation_bookmark("cluster-two"),
    )
    _remove_element(orphan_start)
    _remove_element(orphan_end)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.marker_missing"]
    assert imported.document.blocks == (SemanticParagraph((SemanticText("First [1] then [2]."),)),)


def test_incomplete_orphan_before_a_valid_cluster_still_downgrades_the_whole_paragraph() -> None:
    orphan = _citation("cluster-one", "smith2025", cached_result="[1]")
    valid = _citation("cluster-two", "wang2024", cached_result="[2]")
    document = Document()
    DocxSemanticRenderer(document).render_blocks(
        (
            SemanticParagraph(
                (
                    SemanticText("Orphan "),
                    orphan,
                    SemanticText(" then valid "),
                    valid,
                    SemanticText("."),
                )
            ),
        )
    )
    orphan_start, orphan_end = _bookmark_range(
        document,
        encode_citation_bookmark("cluster-one"),
    )
    paragraph = orphan_start.getparent()
    assert paragraph is not None
    paragraph_children = list(paragraph)
    orphan_field_end = next(
        item
        for root in paragraph_children[
            paragraph_children.index(orphan_start) + 1 : paragraph_children.index(orphan_end)
        ]
        for item in root.iter(qn("w:fldChar"))
        if item.get(qn("w:fldCharType")) == "end"
    )
    _remove_element(orphan_start)
    _remove_element(orphan_end)
    _remove_element(orphan_field_end)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.marker_missing"]
    assert imported.document.blocks == (SemanticParagraph((SemanticText("Orphan [1] then valid [2]."),)),)


def test_split_instruction_orphan_citation_still_has_a_structured_diagnostic() -> None:
    document, _ = _render_citation()
    marker_name = encode_citation_bookmark("cluster-one")
    start, end = _bookmark_range(document, marker_name)
    _remove_element(start)
    _remove_element(end)
    instruction = next(document.element.iter(qn("w:instrText")))
    instruction_run = instruction.getparent()
    assert instruction_run is not None
    instruction.text = " CITA"
    continuation_run = deepcopy(instruction_run)
    continuation = next(continuation_run.iter(qn("w:instrText")))
    continuation.text = "TION smith2025 "
    instruction_run.addnext(continuation_run)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.marker_missing"]
    assert imported.document.blocks == (SemanticParagraph((SemanticText("[1]"),)),)


def test_duplicate_citation_marker_fails_closed_without_importing_either_cluster() -> None:
    document, _ = _render_citation()
    paragraph = document.paragraphs[0]._p
    body = paragraph.getparent()
    assert body is not None
    body.insert(body.index(paragraph) + 1, deepcopy(paragraph))

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == [
        "semantic.docx.citation.marker_invalid",
        "semantic.docx.citation.marker_invalid",
    ]
    assert imported.document.blocks == (
        SemanticParagraph((SemanticText("[1]"),)),
        SemanticParagraph((SemanticText("[1]"),)),
    )


def test_cross_part_duplicate_citation_marker_invalidates_global_proof() -> None:
    document, _ = _render_citation()
    header = document.sections[0].header.paragraphs[0]
    append_zero_width_bookmark(
        header,
        encode_citation_bookmark("cluster-one"),
        "9000",
    )

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.citation.marker_invalid"]
    assert imported.document.blocks == (SemanticParagraph((SemanticText("[1]"),)),)


@pytest.mark.parametrize(
    ("story", "marker_name"),
    [
        ("header", encode_citation_bookmark("header-cluster")),
        ("footer", "_DWC_0"),
    ],
)
def test_package_audit_reports_exact_or_malformed_header_footer_marker(
    story: str,
    marker_name: str,
) -> None:
    document = Document()
    paragraph = getattr(document.sections[0], story).paragraphs[0]
    append_zero_width_bookmark(paragraph, marker_name, "9000")

    imported = DocxSemanticImporter().import_document(document)

    assert [(item.code, item.location) for item in imported.diagnostics] == [
        ("semantic.docx.citation.marker_outside_body", str(paragraph.part.partname))
    ]


def test_package_audit_reports_orphan_marker_inside_a_table_cell() -> None:
    document = Document()
    paragraph = document.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0]
    append_bookmark_start(
        paragraph,
        encode_citation_bookmark("table-cluster"),
        "9000",
    )

    imported = DocxSemanticImporter().import_document(document)

    assert [(item.code, item.location) for item in imported.diagnostics] == [
        ("semantic.docx.citation.marker_outside_body", "/word/document.xml")
    ]


def test_package_audit_reports_marker_in_reopened_generic_word_part(tmp_path) -> None:
    document = Document()
    marker_name = encode_citation_bookmark("footnote-cluster")
    footnotes_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:footnote w:id="1"><w:p>'
        + f'<w:bookmarkStart w:id="9000" w:name="{marker_name}"/>'.encode()
        + b'<w:bookmarkEnd w:id="9000"/>'
        b"</w:p></w:footnote></w:footnotes>"
    )
    footnotes_part = Part(
        PackURI("/word/footnotes.xml"),
        CT.WML_FOOTNOTES,
        footnotes_xml,
        document.part.package,
    )
    document.part.relate_to(footnotes_part, RT.FOOTNOTES)
    source = tmp_path / "citation-footnotes.docx"
    document.save(source)

    imported = DocxSemanticImporter().import_document(Document(source))

    assert [(item.code, item.location) for item in imported.diagnostics] == [
        ("semantic.docx.citation.marker_outside_body", "/word/footnotes.xml")
    ]


def test_package_audit_reports_marker_in_body_paragraph_owned_by_a_caption() -> None:
    document = Document()
    table = SemanticTable(
        row_count=1,
        column_count=1,
        cells=(SemanticTableCell(0, 0, "value"),),
        caption=SemanticCaption(
            kind="table",
            target_id=None,
            cached_number="1",
            label="Table",
            content="Caption-owned paragraph",
        ),
    )
    DocxSemanticRenderer(document).render_table(table)
    append_zero_width_bookmark(
        document.paragraphs[0],
        encode_citation_bookmark("caption-cluster"),
        "9000",
    )

    imported = DocxSemanticImporter().import_document(document)

    assert [(item.code, item.location) for item in imported.diagnostics] == [
        ("semantic.docx.citation.marker_unconsumed", "document/body[0]")
    ]
    assert imported.document.blocks == (table,)


def test_citation_renderer_rejects_reopened_table_cell_without_partial_write(tmp_path) -> None:
    document = Document()
    document.add_table(rows=1, cols=1).cell(0, 0).text = "cell text"
    source = tmp_path / "citation-table-cell-source.docx"
    document.save(source)
    reopened = Document(source)
    paragraph = reopened.tables[0].cell(0, 0).paragraphs[0]
    before = reopened.element.xml

    with pytest.raises(ValueError, match="direct main document body paragraph"):
        DocxSemanticRenderer(reopened).render_citation(paragraph, _citation())

    output = tmp_path / "citation-table-cell-output.docx"
    reopened.save(output)
    final = Document(output)
    assert reopened.element.xml == before
    assert "_DWC_" not in final.element.xml
    assert "CITATION" not in final.element.xml
    assert final.tables[0].cell(0, 0).text == "cell text"


@pytest.mark.parametrize(
    "fragment",
    [
        SemanticBibliographyFragment(entries=()),
        SemanticBibliographyFragment(
            entries=(
                SemanticBibliographyEntry(
                    "smith2025",
                    (SemanticBibliographyRun("Formatted entry."),),
                ),
            )
        ),
    ],
)
def test_bibliography_renderer_rejects_reopened_table_cell_anchor_without_partial_write(
    tmp_path,
    fragment: SemanticBibliographyFragment,
) -> None:
    document = Document()
    document.add_table(rows=1, cols=1).cell(0, 0).text = "bibliography anchor"
    source = tmp_path / "bibliography-table-cell-source.docx"
    document.save(source)
    reopened = Document(source)
    anchor = reopened.tables[0].cell(0, 0).paragraphs[0]
    before = reopened.element.xml

    with pytest.raises(ValueError, match="direct main document body paragraph"):
        DocxSemanticRenderer(reopened).render_bibliography_fragment(
            fragment,
            placeholder_anchor=anchor,
        )

    output = tmp_path / "bibliography-table-cell-output.docx"
    reopened.save(output)
    final = Document(output)
    assert reopened.element.xml == before
    assert "_DWB_" not in final.element.xml
    assert "_DWE_" not in final.element.xml
    assert final.tables[0].cell(0, 0).text == "bibliography anchor"


def test_typed_bibliography_round_trips_with_total_and_entry_markers() -> None:
    document, fragment = _render_bibliography()

    imported = DocxSemanticImporter().import_document(document)
    marker_names = {item.get(qn("w:name")) for item in document.element.iter(qn("w:bookmarkStart"))}

    assert marker_names == {
        BIBLIOGRAPHY_BOOKMARK_NAME,
        encode_bibliography_entry_bookmark("smith2025"),
        encode_bibliography_entry_bookmark("wang2024"),
    }
    assert imported.diagnostics == ()
    assert imported.document == SemanticDocument(
        blocks=(SemanticParagraph((SemanticText("References"),)),),
        bibliography=fragment,
    )
    hyperlink = next(document.element.iter(qn("w:hyperlink")))
    relationship_id = hyperlink.get(qn("r:id"))
    assert relationship_id is not None
    relationship = document.part.rels[relationship_id]
    assert relationship.reltype == RT.HYPERLINK
    assert relationship.is_external
    assert relationship.target_ref == "https://example.org/neutral-documents"
    assert [item.get(qn("w:val")) for item in hyperlink.iter(qn("w:rStyle"))] == ["Hyperlink"]


def test_bibliography_copies_full_paragraph_properties_and_keeps_section_on_last_entry() -> None:
    document = Document()
    anchor = document.add_paragraph("exclusive bibliography anchor")
    anchor.paragraph_format.left_indent = Inches(0.5)
    anchor.paragraph_format.space_after = Pt(7)
    anchor.paragraph_format.keep_with_next = True
    body = document.element.body  # type: ignore[attr-defined]
    section_properties = deepcopy(body.sectPr)
    assert section_properties is not None
    anchor._p.get_or_add_pPr().append(section_properties)
    original_properties = deepcopy(anchor._p.pPr)
    fragment = SemanticBibliographyFragment(
        entries=(
            SemanticBibliographyEntry("first", (SemanticBibliographyRun("First"),)),
            SemanticBibliographyEntry("second", (SemanticBibliographyRun("Second"),)),
        )
    )

    rendered = DocxSemanticRenderer(document).render_bibliography_fragment(
        fragment,
        placeholder_anchor=anchor,
    )

    assert len(rendered) == 2
    assert rendered[0]._p.pPr is not None
    assert rendered[0]._p.pPr.find(qn("w:sectPr")) is None
    final_properties = rendered[1]._p.pPr
    assert final_properties is not None

    def canonical_xml(element: Any) -> bytes:
        return etree.tostring(element, method="c14n", exclusive=True)  # type: ignore[call-overload]

    assert canonical_xml(final_properties) == canonical_xml(original_properties)
    for paragraph in rendered:
        assert paragraph.paragraph_format.left_indent == Inches(0.5)
        assert paragraph.paragraph_format.space_after == Pt(7)
        assert paragraph.paragraph_format.keep_with_next
