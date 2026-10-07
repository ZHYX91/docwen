"""Consume independently authored syntax oracles without sibling dependencies."""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import pytest

from docwen_plugin_markdown.common_utils import parse_raw_md_tables
from docwen_plugin_markdown.document_semantics import analyze_document_semantics
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.number_suite_direct_semantics import analyze_markdown_semantics_v3
from docwen_plugin_markdown.to_docx.notes import decode_internal_note_key, process_md_body_with_notes

pytestmark = pytest.mark.contract

_FIXTURES = Path(__file__).resolve().parents[4] / "tests/fixtures/files/interoperability"
_STRUCTURAL = json.loads((_FIXTURES / "structural-tables.json").read_text(encoding="utf-8"))
_NUMBER_SUITE = json.loads((_FIXTURES / "number-suite.json").read_text(encoding="utf-8"))


def _walk(nodes):
    for node in nodes:
        yield node
        yield from _walk(node.get("children", []))


@pytest.mark.parametrize("fixture", _STRUCTURAL["cases"], ids=lambda fixture: fixture["id"])
def test_structural_tables_independent_corpus(fixture) -> None:
    expected = fixture["expected"]
    analysis = analyze_document_semantics(parse_markdown_text(fixture["source"]), current_v3=True)
    tables = [node for node in _walk(analysis.ast) if node["type"] == "table"]
    raw_tables = parse_raw_md_tables(fixture["source"], structural_tables=True)

    if fixture["class"] == "false_positive":
        assert len(tables) == len(raw_tables) == expected["editable_table_count"]
        assert sum("_structural_table" in table for table in tables) == expected["structural_table_count"]
        return
    if fixture["class"] == "negative":
        if fixture["id"] == "spaced-row-header-boundary":
            # Unlike the editor's invalid-table model, conversion retains literal source.
            assert not tables and not raw_tables
            assert any(node.get("raw") for node in _walk(analysis.ast))
        else:
            assert analysis.has_errors
            code = {
                "merge-missing-anchor": "interop.table.merge_non_rectangular",
                "merge-crosses-role-boundary": "interop.table.merge_role_crossing",
            }[fixture["id"]]
            assert code in [diagnostic.code for diagnostic in analysis.diagnostics]
        return

    assert not analysis.has_errors
    assert len(tables) == len(raw_tables) == expected["structural_table_count"] == 1
    metadata = tables[0]["_document_semantics_table"]
    assert metadata["header_rows"] == expected["header_rows"]
    assert metadata["header_columns"] == expected["row_header_columns"]
    assert metadata["column_count"] == expected["column_count"]
    if "contents" in expected:
        assert raw_tables[0]["all_rows"] == expected["contents"]
    for anchor in expected.get("anchors", []):
        assert any(all(actual[key] == value for key, value in anchor.items()) for actual in metadata["anchors"])
    for cell in expected.get("cells", []):
        actual = next(
            anchor
            for anchor in metadata["anchors"]
            if anchor["row"] == cell["row"] and anchor["column"] == cell["column"]
        )
        assert "".join(node.get("raw", "") for node in _walk(actual["children"])) == cell["content"]
        assert actual["row_span"] == actual["column_span"] == 1


@pytest.mark.parametrize(
    "fixture",
    [case for case in _NUMBER_SUITE["cases"] if case["domain"] == "document"],
    ids=lambda fixture: fixture["id"],
)
def test_number_suite_document_independent_corpus(fixture) -> None:
    analysis = analyze_markdown_semantics_v3(
        fixture["source"],
        input_id="corpus.md",
        consumer_profile="number_suite_direct",
    )
    assert not analysis.has_errors
    projection = analysis.projection
    expected = fixture["expected"]
    kinds = {"figure": "Figure", "table": "Table", "equation": "Equation", "code_block": "Code"}
    if "captions" in expected:
        assert [
            {
                "kind": kinds[item["kind"]],
                "title": item["title"],
                "block_id": item.get("id"),
                "number": int(item["number"]),
            }
            for item in projection["targets"]
            if item["kind"] in kinds
        ] == expected["captions"]
    if "references" in expected:
        assert [
            {"kind": "block", "target": item["target_id"], "alias": item.get("alias")}
            for item in projection["references"]
        ] == expected["references"]
        assert all(item["resolution_status"] == "resolved" for item in projection["references"])


@pytest.mark.parametrize(
    "fixture",
    [case for case in _NUMBER_SUITE["cases"] if case["domain"] == "notes"],
    ids=lambda fixture: fixture["id"],
)
def test_number_suite_notes_independent_corpus(fixture) -> None:
    ast, context = process_md_body_with_notes(fixture["source"])
    expected = fixture["expected"]
    references = expected.get("numbered_references", expected.get("references"))
    counters: Counter[str] = Counter()
    identities = {}
    expected_keys = []
    for reference in references:
        identity = (reference["kind"], reference["id"])
        if identity not in identities:
            counters[identity[0]] += 1
            identities[identity] = str(counters[identity[0]])
        expected_keys.append((identity[0], identities[identity]))
        if "number" in reference:
            assert int(identities[identity]) == reference["number"]
    assert [
        decode_internal_note_key(node["raw"]) for node in _walk(ast) if node["type"] == "footnote_ref"
    ] == expected_keys
    assert len(context._footnote_children) == counters["footnote"]
    assert len(context._endnote_children) == counters["endnote"]
