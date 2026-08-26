"""Focused tests split from test_docx_semantics.py."""

from __future__ import annotations

from ._docx_semantics_support import (
    Document,
    DocxSemanticImporter,
    DocxSemanticRenderer,
    SemanticDocument,
    SemanticDocumentValidationError,
    SemanticParagraph,
    SemanticReference,
    SemanticTable,
    SemanticTableCell,
    SemanticText,
    _force_bookmark_inventory_fallback,
    _insert_simple_seq_before_complex_field,
    _prepare_malformed_table,
    _remove_element,
    _remove_marker_end,
    _render_targeted_table,
    _render_targetless_table,
    _targeted_table,
    _targetless_table,
    _wrap_table_child,
    append_zero_width_bookmark,
    build_docx_bookmark_inventory,
    deepcopy,
    encode_object_bookmark,
    encode_shorthand_bookmark,
    encode_target_bookmark,
    extract_neutral_semantic_caption,
    pytest,
    qn,
)

pytestmark = pytest.mark.contract


def test_targetless_caption_keeps_seq_cache_and_explicit_internal_pairing() -> None:
    document, table = _render_targetless_table()

    imported = DocxSemanticImporter().import_document(document)
    xml = document.element.xml

    assert "SEQ Table" in xml
    assert ">3<" in xml
    assert "_DW_" not in xml
    assert "_DWO_" not in xml
    assert "_DWP_C_" in xml
    assert "_DWP_O_" in xml
    assert imported.diagnostics == ()
    assert imported.document == SemanticDocument(blocks=(table,))


def test_generic_table_import_preserves_a_complete_plain_grid_without_diagnostics() -> None:
    document = Document()
    table = document.add_table(rows=2, cols=2)
    for row, values in enumerate((("A", "B"), ("1", "2"))):
        for column, value in enumerate(values):
            table.cell(row, column).text = value
    document.add_paragraph("Tail paragraph")

    imported = DocxSemanticImporter().import_document(document)

    tables = [block for block in imported.document.blocks if isinstance(block, SemanticTable)]
    assert imported.diagnostics == ()
    assert len(tables) == 1
    assert (tables[0].row_count, tables[0].column_count) == (2, 2)
    assert [cell.text for cell in tables[0].cells] == ["A", "B", "1", "2"]
    assert imported.document.blocks[-1] == SemanticParagraph((SemanticText("Tail paragraph"),))


@pytest.mark.parametrize("wrapper", ["customXml", "sdt"])
@pytest.mark.parametrize("wrapped_kind", ["row", "cell"])
def test_generic_table_import_preserves_wrapped_rows_and_cells(wrapper: str, wrapped_kind: str) -> None:
    document = Document()
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "wrapped"
    if wrapped_kind == "row":
        _wrap_table_child(table._tbl, table.rows[0]._tr, wrapper=wrapper)
    else:
        _wrap_table_child(table.rows[0]._tr, table.cell(0, 0)._tc, wrapper=wrapper)
    document.add_paragraph("Tail paragraph")

    imported = DocxSemanticImporter().import_document(document)

    assert imported.diagnostics == ()
    assert imported.document.blocks == (
        SemanticTable(row_count=1, column_count=1, cells=(SemanticTableCell(0, 0, "wrapped"),)),
        SemanticParagraph((SemanticText("Tail paragraph"),)),
    )


@pytest.mark.parametrize(
    ("malformation", "expected_code"),
    [
        ("rows_missing", "semantic.docx.table.rows_missing"),
        ("cells_missing", "semantic.docx.table.cells_missing"),
        ("grid_span_non_integer", "semantic.docx.table.grid_span_invalid"),
        ("grid_span_zero", "semantic.docx.table.grid_span_invalid"),
        ("grid_span_negative", "semantic.docx.table.grid_span_invalid"),
        ("grid_span_missing", "semantic.docx.table.grid_span_invalid"),
        ("vmerge_orphan", "semantic.docx.table.vmerge_invalid"),
        ("vmerge_mismatched", "semantic.docx.table.vmerge_invalid"),
        ("vmerge_invalid_enum", "semantic.docx.table.vmerge_invalid"),
        ("vmerge_uppercase", "semantic.docx.table.vmerge_invalid"),
        ("grid_before_missing", "semantic.docx.table.grid_offset_invalid"),
        ("grid_after_negative", "semantic.docx.table.grid_offset_invalid"),
        ("grid_before_non_integer", "semantic.docx.table.grid_offset_invalid"),
    ],
)
def test_malformed_generic_table_is_diagnosed_without_losing_following_paragraph(
    malformation: str,
    expected_code: str,
) -> None:
    document = Document()
    _prepare_malformed_table(document, malformation)
    document.add_paragraph("Tail paragraph")

    imported = DocxSemanticImporter().import_document(document)

    assert not any(isinstance(block, SemanticTable) for block in imported.document.blocks)
    assert [(item.level, item.code, item.location) for item in imported.diagnostics] == [
        ("error", expected_code, "document/body[0]")
    ]
    assert imported.document.blocks == (SemanticParagraph((SemanticText("Tail paragraph"),)),)


@pytest.mark.parametrize(
    "malformation",
    ["vmerge_orphan", "vmerge_mismatched", "vmerge_invalid_enum", "vmerge_uppercase"],
)
def test_fallback_import_diagnoses_vertical_merge_without_losing_following_paragraph(malformation: str) -> None:
    document = Document()
    _prepare_malformed_table(document, malformation)
    document.add_paragraph("Tail paragraph")
    _force_bookmark_inventory_fallback(document)

    imported = DocxSemanticImporter().import_document(document)

    assert [(item.level, item.code, item.location) for item in imported.diagnostics] == [
        ("error", "semantic.docx.bookmark.inventory_invalid", "document"),
        ("error", "semantic.docx.table.vmerge_invalid", "document/body[0]"),
    ]
    assert imported.document.blocks == (SemanticParagraph((SemanticText("Tail paragraph"),)),)


def test_twenty_character_target_keeps_all_encoded_bookmark_names_under_word_limit() -> None:
    target_id = "t" * 20

    assert len(encode_target_bookmark(target_id)) == 36
    assert len(encode_object_bookmark(target_id)) == 37
    assert len(encode_shorthand_bookmark(target_id)) == 39


def test_renderer_avoids_sparse_existing_bookmark_ids_across_document_parts() -> None:
    document = Document()
    body = document.add_paragraph("body bookmarks")
    append_zero_width_bookmark(body, "_existing_body_1000", "1000")
    append_zero_width_bookmark(body, "_existing_body_1002", "1002")
    header = document.sections[0].header.paragraphs[0]
    append_zero_width_bookmark(header, "_existing_header_1001", "1001")
    append_zero_width_bookmark(header, "_existing_header_text", "manual")
    footer = document.sections[0].footer.paragraphs[0]
    append_zero_width_bookmark(footer, "_existing_footer_1004", "1004")

    DocxSemanticRenderer(document, bookmark_id_start=1000).render_table(_targetless_table())

    roots = (document.element, header.part.element, footer.part.element)
    starts = [bookmark for root in roots for bookmark in root.iter(qn("w:bookmarkStart"))]
    start_ids = [bookmark.get(qn("w:id")) for bookmark in starts]
    pair_starts = [
        bookmark for bookmark in starts if (bookmark.get(qn("w:name")) or "").startswith(("_DWP_C_", "_DWP_O_"))
    ]
    pair_ids = {bookmark.get(qn("w:id")) for bookmark in pair_starts}
    ends = [bookmark for root in roots for bookmark in root.iter(qn("w:bookmarkEnd"))]

    assert pair_ids == {"1003", "1005"}
    assert "manual" in start_ids
    assert len(start_ids) == len(set(start_ids))
    assert all(sum(end.get(qn("w:id")) == pair_id for end in ends) == 1 for pair_id in pair_ids)


def test_unbalanced_targetless_caption_pair_fails_closed() -> None:
    document, _table = _render_targetless_table()
    caption = document.paragraphs[0]._p
    pair_start = next(
        item for item in caption.iter(qn("w:bookmarkStart")) if (item.get(qn("w:name")) or "").startswith("_DWP_C_")
    )
    pair_id = pair_start.get(qn("w:id"))
    pair_end = next(item for item in caption.iter(qn("w:bookmarkEnd")) if item.get(qn("w:id")) == pair_id)
    _remove_element(pair_end)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.caption_marker_invalid"]
    assert imported.has_errors
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_duplicate_targetless_caption_pair_fails_closed() -> None:
    document, _table = _render_targetless_table()
    caption = document.paragraphs[0]._p
    body = caption.getparent()
    assert body is not None
    body.insert(body.index(caption) + 1, deepcopy(caption))

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.caption_duplicate"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_targetless_pair_with_shorthand_marker_fails_closed() -> None:
    document, _table = _render_targetless_table()
    append_zero_width_bookmark(
        document.paragraphs[0],
        encode_shorthand_bookmark("tbl-other"),
        "9000",
    )

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.shorthand_marker_invalid"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_unbalanced_targetless_object_pair_fails_closed() -> None:
    document, _table = _render_targetless_table()
    table_element = document.tables[0]._tbl
    pair_start = next(
        item
        for item in table_element.iter(qn("w:bookmarkStart"))
        if (item.get(qn("w:name")) or "").startswith("_DWP_O_")
    )
    pair_id = pair_start.get(qn("w:id"))
    pair_end = next(item for item in table_element.iter(qn("w:bookmarkEnd")) if item.get(qn("w:id")) == pair_id)
    _remove_element(pair_end)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.object_marker_invalid"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_duplicate_targetless_object_pair_fails_closed() -> None:
    document, _table = _render_targetless_table()
    table_element = document.tables[0]._tbl
    body = table_element.getparent()
    assert body is not None
    body.insert(body.index(table_element) + 1, deepcopy(table_element))

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.object_duplicate"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_orphan_targetless_caption_pair_fails_closed() -> None:
    document, _table = _render_targetless_table()
    table_element = document.tables[0]._tbl
    _remove_element(table_element)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.object_missing"]
    assert imported.has_errors


def test_orphan_targetless_object_pair_fails_closed() -> None:
    document, _table = _render_targetless_table()
    caption = document.paragraphs[0]._p
    _remove_element(caption)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.caption_missing"]
    assert imported.has_errors
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_unbalanced_target_caption_bookmark_fails_closed() -> None:
    document, _table = _render_targeted_table()
    _remove_marker_end(document.paragraphs[0]._p, prefix="_DW_")

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.caption_target_marker_invalid"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_target_caption_binds_to_its_table_without_adjacency() -> None:
    document, table = _render_targeted_table()
    caption_element = document.paragraphs[0]._p
    body = caption_element.getparent()
    assert body is not None
    separator = document.add_paragraph("ordinary paragraph between caption and table")._p
    body.remove(separator)
    body.insert(body.index(caption_element) + 1, separator)

    imported = DocxSemanticImporter().import_document(document)

    assert imported.diagnostics == ()
    assert table in imported.document.blocks


@pytest.mark.parametrize("localized_identifier", ["表格", "Tabelle"])
def test_localized_seq_identifier_uses_proven_target_table_kind(localized_identifier: str) -> None:
    document = Document()
    table = _targeted_table()
    semantic_document = SemanticDocument(
        blocks=(
            table,
            SemanticParagraph((SemanticReference("tbl-sales", "7"),)),
        )
    )
    DocxSemanticRenderer(document).render_blocks(semantic_document.blocks)
    instruction = next(
        item for item in document.element.iter(qn("w:instrText")) if (item.text or "").strip().startswith("SEQ Table")
    )
    instruction.text = (instruction.text or "").replace("Table", localized_identifier)

    imported = DocxSemanticImporter().import_document(document)

    assert imported.diagnostics == ()
    assert imported.document == semantic_document


@pytest.mark.parametrize(
    ("known_identifier", "expected_kind"),
    [("Figure", "figure"), ("Table", "table"), ("Equation", "equation"), ("Listing", "listing")],
)
def test_known_seq_after_unknown_precedes_structural_table_fallback(
    known_identifier: str,
    expected_kind: str,
) -> None:
    document, _table = _render_targeted_table()
    paragraph = document.paragraphs[0]
    instruction = next(
        item for item in paragraph._p.iter(qn("w:instrText")) if (item.text or "").strip().startswith("SEQ Table")
    )
    instruction.text = (instruction.text or "").replace("Table", known_identifier)
    _insert_simple_seq_before_complex_field(paragraph, identifier="Unknown", cached_result="999")

    extracted = extract_neutral_semantic_caption(
        paragraph._p,
        bookmark_inventory=build_docx_bookmark_inventory(document),
        proven_object_kind="table",
    )

    assert extracted is not None
    assert extracted.kind == expected_kind
    assert extracted.cached_number == "7"


def test_multiple_known_seq_identifiers_keep_first_known_field_order() -> None:
    document, _table = _render_targeted_table()
    paragraph = document.paragraphs[0]
    _insert_simple_seq_before_complex_field(paragraph, identifier="Figure", cached_result="2")

    extracted = extract_neutral_semantic_caption(
        paragraph._p,
        bookmark_inventory=build_docx_bookmark_inventory(document),
        proven_object_kind="table",
    )

    assert extracted is not None
    assert extracted.kind == "figure"
    assert extracted.cached_number == "2"


@pytest.mark.parametrize("localized_identifier", ["表格", "Tabelle"])
def test_localized_seq_identifier_uses_proven_targetless_table_pairing(
    localized_identifier: str,
    tmp_path,
) -> None:
    document, table = _render_targetless_table()
    instruction = next(
        item for item in document.element.iter(qn("w:instrText")) if (item.text or "").strip().startswith("SEQ Table")
    )
    instruction.text = (instruction.text or "").replace("Table", localized_identifier)
    output = tmp_path / f"targetless-{localized_identifier}.docx"
    document.save(output)

    imported = DocxSemanticImporter().import_document(Document(output))

    assert imported.diagnostics == ()
    assert imported.document == SemanticDocument(blocks=(table,))


def test_multiple_unknown_seq_identifiers_fail_closed_despite_proven_targetless_pairing() -> None:
    document, _table = _render_targetless_table()
    paragraph = document.paragraphs[0]
    instruction = next(
        item for item in paragraph._p.iter(qn("w:instrText")) if (item.text or "").strip().startswith("SEQ Table")
    )
    instruction.text = (instruction.text or "").replace("Table", "表格")
    _insert_simple_seq_before_complex_field(paragraph, identifier="Tabelle", cached_result="99")

    imported = DocxSemanticImporter().import_document(document)
    imported_tables = [block for block in imported.document.blocks if isinstance(block, SemanticTable)]

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.caption_missing"]
    assert len(imported_tables) == 1
    assert imported_tables[0].caption is None


def test_localized_targetless_seq_requires_balanced_object_pairing_marker() -> None:
    document, _table = _render_targetless_table()
    instruction = next(
        item for item in document.element.iter(qn("w:instrText")) if (item.text or "").strip().startswith("SEQ Table")
    )
    instruction.text = (instruction.text or "").replace("Table", "表格")
    _remove_marker_end(document.tables[0]._tbl, prefix="_DWP_O_")

    imported = DocxSemanticImporter().import_document(document)
    imported_tables = [block for block in imported.document.blocks if isinstance(block, SemanticTable)]

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.object_marker_invalid"]
    assert len(imported_tables) == 1
    assert imported_tables[0].caption is None


def test_localized_targetless_seq_rejects_conflicting_object_target_marker() -> None:
    document, _table = _render_targetless_table()
    instruction = next(
        item for item in document.element.iter(qn("w:instrText")) if (item.text or "").strip().startswith("SEQ Table")
    )
    instruction.text = (instruction.text or "").replace("Table", "表格")
    DocxSemanticRenderer(document, bookmark_id_start=2000).bind_object_target(document.tables[0], "tbl-conflict")

    imported = DocxSemanticImporter().import_document(document)
    imported_tables = [block for block in imported.document.blocks if isinstance(block, SemanticTable)]

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.object_marker_conflict"]
    assert len(imported_tables) == 1
    assert imported_tables[0].caption is None


def test_localized_seq_identifier_requires_matching_physical_table_target() -> None:
    document = Document()
    semantic_document = SemanticDocument(
        blocks=(
            _targeted_table(),
            SemanticParagraph((SemanticReference("tbl-sales", "7"),)),
        )
    )
    DocxSemanticRenderer(document).render_blocks(semantic_document.blocks)
    instruction = next(
        item for item in document.element.iter(qn("w:instrText")) if (item.text or "").strip().startswith("SEQ Table")
    )
    instruction.text = (instruction.text or "").replace("Table", "表格")
    object_start = next(
        item
        for item in document.tables[0]._tbl.iter(qn("w:bookmarkStart"))
        if (item.get(qn("w:name")) or "").startswith("_DWO_")
    )
    object_id = object_start.get(qn("w:id"))
    object_end = next(
        item for item in document.tables[0]._tbl.iter(qn("w:bookmarkEnd")) if item.get(qn("w:id")) == object_id
    )
    object_start_parent = object_start.getparent()
    object_end_parent = object_end.getparent()
    assert object_start_parent is not None
    assert object_end_parent is not None
    object_start_parent.remove(object_start)
    object_end_parent.remove(object_end)

    imported = DocxSemanticImporter().import_document(document)
    imported_tables = [block for block in imported.document.blocks if isinstance(block, SemanticTable)]

    assert [item.code for item in imported.diagnostics] == ["semantic.reference.target_missing"]
    assert len(imported_tables) == 1
    assert imported_tables[0].caption is None


def test_reference_to_unbalanced_target_imports_only_its_cached_text() -> None:
    document = Document()
    renderer = DocxSemanticRenderer(document)
    renderer.render_blocks(
        (
            _targeted_table(),
            SemanticParagraph((SemanticReference("tbl-sales", "7"),)),
        )
    )
    _remove_marker_end(document.paragraphs[0]._p, prefix="_DW_")

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.caption_target_marker_invalid"]
    assert SemanticParagraph((SemanticText("7"),)) in imported.document.blocks
    assert all(
        not isinstance(inline, SemanticReference)
        for block in imported.document.blocks
        if isinstance(block, SemanticParagraph)
        for inline in block.inlines
    )


def test_unbalanced_target_object_bookmark_fails_closed() -> None:
    document, _table = _render_targeted_table()
    _remove_marker_end(document.tables[0]._tbl, prefix="_DWO_")

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.object_target_marker_invalid"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_duplicate_target_caption_does_not_bind_either_caption() -> None:
    document, _table = _render_targeted_table()
    caption = document.paragraphs[0]._p
    body = caption.getparent()
    assert body is not None
    body.insert(body.index(caption) + 1, deepcopy(caption))

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.caption.target_duplicate"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_duplicate_target_table_never_reuses_one_caption() -> None:
    document, _table = _render_targeted_table()
    table_element = document.tables[0]._tbl
    body = table_element.getparent()
    assert body is not None
    body.insert(body.index(table_element) + 1, deepcopy(table_element))

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.object_target_duplicate"]
    semantic_tables = [block for block in imported.document.blocks if isinstance(block, SemanticTable)]
    assert len(semantic_tables) == 2
    assert all(table.caption is None for table in semantic_tables)


def test_existing_same_name_invalidates_target_proof() -> None:
    document, _table = _render_targeted_table()
    unrelated = document.add_paragraph("unrelated")
    append_zero_width_bookmark(unrelated, encode_target_bookmark("tbl-sales"), "9000")

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.caption_target_marker_invalid"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_unrelated_bookmark_reusing_pair_id_invalidates_pair_proof() -> None:
    document, _table = _render_targetless_table()
    pair_start = next(
        item
        for item in document.paragraphs[0]._p.iter(qn("w:bookmarkStart"))
        if (item.get(qn("w:name")) or "").startswith("_DWP_C_")
    )
    unrelated = document.add_paragraph("unrelated")
    pair_id = pair_start.get(qn("w:id"))
    assert pair_id is not None
    append_zero_width_bookmark(unrelated, "_unrelated", pair_id)

    imported = DocxSemanticImporter().import_document(document)

    assert [item.code for item in imported.diagnostics] == ["semantic.docx.binding.caption_marker_invalid"]
    assert all(not isinstance(block, SemanticTable) or block.caption is None for block in imported.document.blocks)


def test_renderer_rejects_existing_semantic_bookmark_name_without_partial_write() -> None:
    document = Document()
    unrelated = document.add_paragraph("unrelated")
    append_zero_width_bookmark(unrelated, encode_target_bookmark("tbl-sales"), "12")
    table = document.add_table(rows=1, cols=1)
    renderer = DocxSemanticRenderer(document)

    with pytest.raises(SemanticDocumentValidationError) as error:
        renderer.render_caption_for_table(table, _targeted_table().caption)  # type: ignore[arg-type]

    assert [item.code for item in error.value.diagnostics] == ["semantic.docx.bookmark.name_conflict"]
    assert "SEQ Table" not in document.element.xml
    assert "_DWO_" not in document.element.xml


def test_atomic_table_caption_rejects_a_table_from_another_document() -> None:
    document = Document()
    other_document = Document()
    foreign_table = other_document.add_table(rows=1, cols=1)
    document_before = document.element.xml
    table_before = foreign_table._tbl.xml

    with pytest.raises(ValueError, match="must belong to the renderer document part"):
        DocxSemanticRenderer(document).render_caption_for_table(
            foreign_table,
            _targetless_table().caption,  # type: ignore[arg-type]
        )

    assert document.element.xml == document_before
    assert foreign_table._tbl.xml == table_before
