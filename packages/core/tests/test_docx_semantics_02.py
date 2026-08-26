"""Focused tests split from test_docx_semantics.py."""

from __future__ import annotations

from ._docx_semantics_support import (
    BIBLIOGRAPHY_BOOKMARK_NAME,
    BOOKMARK_ID_MAX,
    CT,
    RT,
    Document,
    DocxSemanticImporter,
    DocxSemanticRenderer,
    PackURI,
    Part,
    SemanticBibliographyEntry,
    SemanticBibliographyFragment,
    SemanticBibliographyRun,
    SemanticDocumentValidationError,
    SemanticParagraph,
    SemanticText,
    _targeted_table,
    _targetless_table,
    append_bookmark_end,
    append_zero_width_bookmark,
    build_docx_bookmark_inventory,
    encode_object_bookmark,
    encode_object_pairing_bookmark,
    encode_shorthand_bookmark,
    encode_target_bookmark,
    pytest,
    qn,
)

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("method_name", "binding_value"),
    [("bind_object_target", "tbl-sales"), ("bind_object_pairing", "1")],
)
def test_public_table_binding_helpers_reject_a_table_from_another_document(
    method_name: str,
    binding_value: str,
) -> None:
    document = Document()
    other_document = Document()
    foreign_table = other_document.add_table(rows=1, cols=1)
    document_before = document.element.xml
    table_before = foreign_table._tbl.xml

    with pytest.raises(ValueError, match="must belong to the renderer document part"):
        getattr(DocxSemanticRenderer(document), method_name)(foreign_table, binding_value)

    assert document.element.xml == document_before
    assert foreign_table._tbl.xml == table_before


@pytest.mark.parametrize(
    ("method_name", "binding_value"),
    [("bind_object_target", "tbl-sales"), ("bind_object_pairing", "1")],
)
def test_public_table_binding_helpers_reject_an_empty_table(
    method_name: str,
    binding_value: str,
) -> None:
    document = Document()
    empty_table = document.add_table(rows=0, cols=0)

    with pytest.raises(ValueError, match="only to a non-empty table"):
        getattr(DocxSemanticRenderer(document), method_name)(empty_table, binding_value)

    assert "_DWO_" not in document.element.xml
    assert "_DWP_O_" not in document.element.xml


def test_atomic_table_caption_preflights_object_and_shorthand_names() -> None:
    document = Document()
    unrelated = document.add_paragraph("unrelated")
    append_zero_width_bookmark(unrelated, encode_object_bookmark("tbl-sales"), "12")
    append_zero_width_bookmark(unrelated, encode_shorthand_bookmark("tbl-sales"), "13")
    table = document.add_table(rows=1, cols=1)

    with pytest.raises(SemanticDocumentValidationError) as error:
        DocxSemanticRenderer(document).render_caption_for_table(
            table,
            _targeted_table().caption,  # type: ignore[arg-type]
            source_form="shorthand",
        )

    assert [item.code for item in error.value.diagnostics] == ["semantic.docx.bookmark.name_conflict"]
    assert "SEQ Table" not in document.element.xml
    assert encode_target_bookmark("tbl-sales") not in document.element.xml


def test_explicit_pair_and_bibliography_name_conflicts_are_rejected() -> None:
    document = Document()
    unrelated = document.add_paragraph("unrelated")
    append_zero_width_bookmark(unrelated, encode_object_pairing_bookmark("1"), "12")
    append_zero_width_bookmark(unrelated, BIBLIOGRAPHY_BOOKMARK_NAME, "13")
    table = document.add_table(rows=1, cols=1)
    renderer = DocxSemanticRenderer(document)

    with pytest.raises(SemanticDocumentValidationError):
        renderer.bind_object_pairing(table, "1")

    anchor = document.add_paragraph("bibliography anchor")
    with pytest.raises(SemanticDocumentValidationError):
        renderer.render_bibliography_fragment(
            SemanticBibliographyFragment(
                entries=(SemanticBibliographyEntry("entry", (SemanticBibliographyRun("Entry"),)),)
            ),
            placeholder_anchor=anchor,
        )

    assert anchor._element.getparent() is not None


def test_multiple_renderer_instances_rescan_names_and_ids() -> None:
    document = Document()
    first_table = document.add_table(rows=1, cols=1)
    second_table = document.add_table(rows=1, cols=1)
    first_renderer = DocxSemanticRenderer(document, bookmark_id_start=1000)
    second_renderer = DocxSemanticRenderer(document, bookmark_id_start=1000)

    first_renderer.render_caption_for_table(first_table, _targetless_table().caption)  # type: ignore[arg-type]
    second_renderer.render_caption_for_table(second_table, _targetless_table().caption)  # type: ignore[arg-type]

    inventory = build_docx_bookmark_inventory(document)
    assert len(inventory.starts) == 4
    assert len({item.id_key for item in inventory.starts}) == 4
    assert len({item.name.casefold() for item in inventory.starts if item.name is not None}) == 4


@pytest.mark.parametrize("bookmark_id_start", [True, 1.5, "1000", None])
def test_bookmark_id_start_rejects_non_integer_values(bookmark_id_start) -> None:
    with pytest.raises(TypeError, match="must be an integer"):
        DocxSemanticRenderer(Document(), bookmark_id_start=bookmark_id_start)


@pytest.mark.parametrize("bookmark_id_start", [-1, BOOKMARK_ID_MAX + 1])
def test_bookmark_id_start_rejects_values_outside_portable_range(bookmark_id_start: int) -> None:
    with pytest.raises(ValueError, match="must be between"):
        DocxSemanticRenderer(Document(), bookmark_id_start=bookmark_id_start)


def test_bookmark_allocator_avoids_orphan_end_and_decimal_lexical_aliases() -> None:
    document = Document()
    paragraph = document.add_paragraph("existing")
    append_bookmark_end(paragraph, "1000")
    append_zero_width_bookmark(paragraph, "_leading_zero", "01001")

    DocxSemanticRenderer(document, bookmark_id_start=1000).render_table(_targetless_table())

    pair_ids = {
        item.get(qn("w:id"))
        for item in document.element.iter(qn("w:bookmarkStart"))
        if (item.get(qn("w:name")) or "").startswith(("_DWP_C_", "_DWP_O_"))
    }
    assert pair_ids == {"1002", "1003"}
    assert sum(item.get(qn("w:id")) == "1000" for item in document.element.iter(qn("w:bookmarkEnd"))) == 1


def test_bookmark_allocator_treats_leading_zero_decimal_id_as_occupied() -> None:
    document = Document()
    paragraph = document.add_paragraph("existing")
    append_zero_width_bookmark(paragraph, "_leading_zero", "01000")

    DocxSemanticRenderer(document, bookmark_id_start=1000).render_table(_targetless_table())

    pair_ids = {
        item.get(qn("w:id"))
        for item in document.element.iter(qn("w:bookmarkStart"))
        if (item.get(qn("w:name")) or "").startswith(("_DWP_C_", "_DWP_O_"))
    }
    assert pair_ids == {"1001", "1002"}


def test_bookmark_allocator_overflow_fails_before_writing_semantic_markers() -> None:
    document = Document()
    table = document.add_table(rows=1, cols=1)
    renderer = DocxSemanticRenderer(document, bookmark_id_start=BOOKMARK_ID_MAX)

    with pytest.raises(OverflowError, match="no portable DOCX bookmark IDs remain"):
        renderer.render_caption_for_table(table, _targetless_table().caption)  # type: ignore[arg-type]

    assert "_DWP_" not in document.element.xml


def test_generic_footnotes_part_is_inventoried_after_real_save_reopen(tmp_path) -> None:
    document = Document()
    footnotes_xml = (
        b'\xef\xbb\xbf<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:footnote w:id="1"><w:p>'
        b'<w:bookmarkStart w:id="1000" w:name="_footnote_existing"/>'
        b'<w:bookmarkEnd w:id="1000"/>'
        b"</w:p></w:footnote></w:footnotes>"
    )
    footnotes_part = Part(
        PackURI("/word/footnotes.xml"),
        CT.WML_FOOTNOTES,
        footnotes_xml,
        document.part.package,
    )
    document.part.relate_to(footnotes_part, RT.FOOTNOTES)
    seed = tmp_path / "footnotes-seed.docx"
    document.save(seed)

    reopened = Document(seed)
    DocxSemanticRenderer(reopened, bookmark_id_start=1000).render_table(_targetless_table())
    output = tmp_path / "footnotes-rendered.docx"
    reopened.save(output)
    final = Document(output)
    inventory = build_docx_bookmark_inventory(final)
    pair_ids = {
        item.raw_id
        for item in inventory.starts
        if item.name is not None and item.name.startswith(("_DWP_C_", "_DWP_O_"))
    }

    assert pair_ids == {"1001", "1002"}
    assert len(inventory.starts_with_id(("numeric", 1000))) == 1


def test_inventory_does_not_parse_reachable_non_xml_word_part() -> None:
    document = Document()
    binary_part = Part(
        PackURI("/word/media/opaque.bin"),
        "application/octet-stream",
        b"not xml and intentionally opaque",
        document.part.package,
    )
    document.part.relate_to(binary_part, "urn:docwen:test:opaque")

    inventory = build_docx_bookmark_inventory(document)

    assert inventory.starts == ()
    assert inventory.ends == ()


def test_malformed_reachable_word_xml_part_fails_closed_without_leaking_content() -> None:
    document = Document()
    document.add_paragraph("ordinary content")
    sentinel = "private-part-content"
    malformed_part = Part(
        PackURI("/word/comments.xml"),
        CT.WML_COMMENTS,
        sentinel.encode("ascii"),
        document.part.package,
    )
    document.part.relate_to(malformed_part, RT.COMMENTS)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.bookmark.inventory_invalid"]
    assert imported.document.blocks == (SemanticParagraph((SemanticText("ordinary content"),)),)
    assert sentinel not in imported.diagnostics[0].message


def test_empty_bibliography_deletes_only_its_explicit_anchor() -> None:
    document = Document()
    heading = document.add_paragraph("References")
    anchor = document.add_paragraph("exclusive bibliography anchor")
    tail = document.add_paragraph("After bibliography")

    rendered = DocxSemanticRenderer(document).render_bibliography_fragment(
        SemanticBibliographyFragment(entries=()),
        placeholder_anchor=anchor,
    )

    assert rendered == ()
    assert [paragraph.text for paragraph in document.paragraphs] == [heading.text, tail.text]


def test_unbalanced_bibliography_boundary_fails_closed() -> None:
    document = Document()
    document.add_paragraph("References")
    anchor = document.add_paragraph("exclusive bibliography anchor")
    DocxSemanticRenderer(document).render_bibliography_fragment(
        SemanticBibliographyFragment(
            entries=(
                SemanticBibliographyEntry("first", (SemanticBibliographyRun("First entry."),)),
                SemanticBibliographyEntry("second", (SemanticBibliographyRun("Second entry."),)),
            )
        ),
        placeholder_anchor=anchor,
    )
    boundary_start = next(
        item
        for item in document.element.iter(qn("w:bookmarkStart"))
        if item.get(qn("w:name")) == BIBLIOGRAPHY_BOOKMARK_NAME
    )
    boundary_id = boundary_start.get(qn("w:id"))
    boundary_end = next(
        item for item in document.element.iter(qn("w:bookmarkEnd")) if item.get(qn("w:id")) == boundary_id
    )
    boundary_parent = boundary_end.getparent()
    assert boundary_parent is not None
    boundary_parent.remove(boundary_end)

    imported = DocxSemanticImporter().import_document(document)

    assert imported.document.bibliography is None
    assert [(item.level, item.code, item.location) for item in imported.diagnostics] == [
        (
            "error",
            "semantic.docx.bibliography.boundary_unbalanced",
            "document/body[1]",
        )
    ]
