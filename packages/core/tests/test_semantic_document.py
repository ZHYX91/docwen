"""Focused tests for the provider-neutral semantic document model."""

from __future__ import annotations

import pytest

from docwen_core.models.semantic_document import (
    SemanticBibliographyEntry,
    SemanticBibliographyFragment,
    SemanticBibliographyRun,
    SemanticCaption,
    SemanticCitationCluster,
    SemanticCitationItem,
    SemanticDocument,
    SemanticParagraph,
    SemanticReference,
    SemanticTable,
    SemanticTableCell,
    SemanticText,
    derive_table_header_shape,
    validate_semantic_document,
)

pytestmark = pytest.mark.unit


def _neutral_table() -> SemanticTable:
    # Four total rows are necessary for all four roles: the first two rows
    # form the requested 2x4 column-header region and the final two carry
    # row-header/data cells.
    return SemanticTable(
        row_count=4,
        column_count=4,
        repeat_header="always",
        caption=SemanticCaption(
            kind="table",
            target_id="tbl-sales",
            cached_number="7",
            label="Table",
            content="Sales channels",
        ),
        cells=(
            SemanticTableCell(0, 0, "Region", "corner_header", row_span=2),
            SemanticTableCell(0, 1, "Sales", "column_header", column_span=2),
            SemanticTableCell(0, 3, "Total", "column_header"),
            SemanticTableCell(1, 1, "Online", "column_header"),
            SemanticTableCell(1, 2, "Retail", "column_header"),
            SemanticTableCell(1, 3, "Combined", "column_header"),
            SemanticTableCell(2, 0, "North", "row_header", row_span=2),
            SemanticTableCell(2, 1, "10", "data", column_span=2),
            SemanticTableCell(2, 3, "22", "data"),
            SemanticTableCell(3, 1, "8", "data"),
            SemanticTableCell(3, 2, "9", "data"),
            SemanticTableCell(3, 3, "17", "data"),
        ),
    )


def test_explicit_anchor_grid_and_target_bound_reference_are_valid() -> None:
    table = _neutral_table()
    document = SemanticDocument(
        blocks=(
            table,
            SemanticParagraph(
                (
                    SemanticText("See table "),
                    SemanticReference("tbl-sales", "7"),
                    SemanticText("."),
                )
            ),
        )
    )

    assert validate_semantic_document(document) == ()
    assert derive_table_header_shape(table) == (2, 1)


def test_incomplete_grid_and_missing_reference_have_structured_diagnostics() -> None:
    table = SemanticTable(
        row_count=1,
        column_count=2,
        cells=(SemanticTableCell(0, 0, "only", "data"),),
    )
    document = SemanticDocument(
        blocks=(
            table,
            SemanticParagraph((SemanticReference("tbl-missing", "1"),)),
        )
    )

    assert [(item.code, item.location) for item in validate_semantic_document(document)] == [
        ("semantic.table.grid_incomplete", "blocks[0].cell[0,1]"),
        ("semantic.reference.target_missing", "blocks[1].inlines[0]"),
    ]


def test_span_cannot_cross_explicit_role_boundary() -> None:
    table = SemanticTable(
        row_count=2,
        column_count=2,
        cells=(
            SemanticTableCell(0, 0, "bad", "column_header", row_span=2),
            SemanticTableCell(0, 1, "header", "column_header"),
            SemanticTableCell(1, 1, "data", "data"),
        ),
    )

    diagnostics = validate_semantic_document(SemanticDocument(blocks=(table,)))

    assert [item.code for item in diagnostics] == ["semantic.table.role_boundary_invalid"]


def test_portable_target_id_accepts_twenty_ascii_characters() -> None:
    target_id = "t" * 20
    table = SemanticTable(
        row_count=1,
        column_count=1,
        cells=(SemanticTableCell(0, 0, "value"),),
        caption=SemanticCaption("table", target_id, "1", "Table", "Portable target"),
    )
    document = SemanticDocument(
        blocks=(table, SemanticParagraph((SemanticReference(target_id, "1"),))),
    )

    assert validate_semantic_document(document) == ()


def test_portable_target_id_rejects_twenty_one_characters_for_caption_and_reference() -> None:
    target_id = "t" * 21
    table = SemanticTable(
        row_count=1,
        column_count=1,
        cells=(SemanticTableCell(0, 0, "value"),),
        caption=SemanticCaption("table", target_id, "1", "Table", "Too long"),
    )
    document = SemanticDocument(
        blocks=(table, SemanticParagraph((SemanticReference(target_id, "1"),))),
    )

    assert [item.code for item in validate_semantic_document(document)] == [
        "semantic.caption.target_id_invalid",
        "semantic.reference.target_id_invalid",
    ]


def test_portable_target_id_rejects_non_ascii_for_caption_and_reference() -> None:
    target_id = "表格"
    table = SemanticTable(
        row_count=1,
        column_count=1,
        cells=(SemanticTableCell(0, 0, "value"),),
        caption=SemanticCaption("table", target_id, "1", "Table", "Non-ASCII"),
    )
    document = SemanticDocument(
        blocks=(table, SemanticParagraph((SemanticReference(target_id, "1"),))),
    )

    assert [item.code for item in validate_semantic_document(document)] == [
        "semantic.caption.target_id_invalid",
        "semantic.reference.target_id_invalid",
    ]


def test_portable_citation_and_bibliography_ids_accept_twenty_ascii_characters() -> None:
    identifier = "a" * 20
    document = SemanticDocument(
        blocks=(
            SemanticParagraph(
                (
                    SemanticCitationCluster(
                        cluster_id=identifier,
                        items=(SemanticCitationItem(identifier),),
                        cached_result="[1]",
                    ),
                )
            ),
        ),
        bibliography=SemanticBibliographyFragment(
            entries=(SemanticBibliographyEntry(identifier, (SemanticBibliographyRun("Formatted entry."),)),)
        ),
    )

    assert validate_semantic_document(document) == ()


@pytest.mark.parametrize("identifier", ["", "a" * 21, "1starts-with-digit", "citation-表"])
def test_citation_and_bibliography_ids_reject_nonportable_values(identifier: str) -> None:
    document = SemanticDocument(
        blocks=(
            SemanticParagraph(
                (
                    SemanticCitationCluster(
                        cluster_id=identifier,
                        items=(SemanticCitationItem(identifier),),
                        cached_result="[1]",
                    ),
                )
            ),
        ),
        bibliography=SemanticBibliographyFragment(
            entries=(SemanticBibliographyEntry(identifier, (SemanticBibliographyRun("Formatted entry."),)),)
        ),
    )

    assert [item.code for item in validate_semantic_document(document)] == [
        "semantic.citation.cluster_id_invalid",
        "semantic.citation.item_id_invalid",
        "semantic.bibliography.item_id_invalid",
    ]


def test_duplicate_clusters_and_bibliography_items_are_rejected() -> None:
    cluster = SemanticCitationCluster(
        cluster_id="cluster-one",
        items=(SemanticCitationItem("item-one"),),
        cached_result="[1]",
    )
    document = SemanticDocument(
        blocks=(SemanticParagraph((cluster, cluster)),),
        bibliography=SemanticBibliographyFragment(
            entries=(
                SemanticBibliographyEntry("item-one", (SemanticBibliographyRun("First."),)),
                SemanticBibliographyEntry("item-one", (SemanticBibliographyRun("Duplicate."),)),
            )
        ),
    )

    assert [item.code for item in validate_semantic_document(document)] == [
        "semantic.citation.cluster_id_duplicate",
        "semantic.bibliography.item_id_duplicate",
    ]


def test_empty_citation_items_cache_and_bibliography_content_are_rejected() -> None:
    document = SemanticDocument(
        blocks=(
            SemanticParagraph(
                (
                    SemanticCitationCluster(
                        cluster_id="cluster-one",
                        items=(),
                        cached_result="   ",
                    ),
                )
            ),
        ),
        bibliography=SemanticBibliographyFragment(
            entries=(SemanticBibliographyEntry("item-one", (SemanticBibliographyRun("\t"),)),)
        ),
    )

    assert [item.code for item in validate_semantic_document(document)] == [
        "semantic.citation.items_empty",
        "semantic.citation.cached_result_empty",
        "semantic.bibliography.entry_empty",
    ]


def test_bibliography_requires_at_least_one_run() -> None:
    document = SemanticDocument(
        blocks=(),
        bibliography=SemanticBibliographyFragment(entries=(SemanticBibliographyEntry("item-one", ()),)),
    )

    assert [item.code for item in validate_semantic_document(document)] == ["semantic.bibliography.entry_empty"]


@pytest.mark.parametrize(
    ("run", "expected_code"),
    [
        (SemanticBibliographyRun(""), "semantic.bibliography.run_text_empty"),
        (SemanticBibliographyRun("bad\x00text"), "semantic.bibliography.run_text_empty"),
        (SemanticBibliographyRun("bad\ud800text"), "semantic.bibliography.run_text_empty"),
        (
            SemanticBibliographyRun("entry", bold=1),  # type: ignore[arg-type]
            "semantic.bibliography.run_bold_invalid",
        ),
        (
            SemanticBibliographyRun("entry", italic=1),  # type: ignore[arg-type]
            "semantic.bibliography.run_italic_invalid",
        ),
        (
            SemanticBibliographyRun("entry", href="mailto:test@example.org"),
            "semantic.bibliography.run_href_invalid",
        ),
        (
            SemanticBibliographyRun("entry", href="/relative"),
            "semantic.bibliography.run_href_invalid",
        ),
    ],
)
def test_bibliography_run_contract_is_fail_closed(
    run: SemanticBibliographyRun,
    expected_code: str,
) -> None:
    document = SemanticDocument(
        blocks=(),
        bibliography=SemanticBibliographyFragment(entries=(SemanticBibliographyEntry("item-one", (run,)),)),
    )

    assert expected_code in {item.code for item in validate_semantic_document(document)}


def test_bibliography_accepts_ordered_rich_http_runs() -> None:
    document = SemanticDocument(
        blocks=(),
        bibliography=SemanticBibliographyFragment(
            entries=(
                SemanticBibliographyEntry(
                    "item-one",
                    (
                        SemanticBibliographyRun("Author. ", bold=True),
                        SemanticBibliographyRun(
                            "Title",
                            italic=True,
                            href="https://example.org/title",
                        ),
                        SemanticBibliographyRun("."),
                    ),
                ),
            )
        ),
    )

    assert validate_semantic_document(document) == ()
