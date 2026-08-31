"""Cross-type caption ownership contracts for the Markdown plugin."""

from __future__ import annotations

import pytest

from docwen_plugin_markdown.document_semantics_v3 import analyze_markdown_semantics_v3
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.runtime_semantics_v3 import (
    apply_runtime_semantics_v3,
    prepare_runtime_semantics_v3,
)

pytestmark = pytest.mark.contract

_CAPTIONS = {
    "figure": "Figure: Composite",
    "table": "Table: Composite",
    "equation": "Equation: Composite",
    "code_block": "Code: Composite",
}
_OBJECTS = {
    "figure": "![image](image.png)",
    "table": "| A |\n|---|\n| 1 |",
    "equation": "$$x=1$$",
    "code_block": "```text\nx\n```",
}


@pytest.mark.parametrize("caption_kind", tuple(_CAPTIONS))
@pytest.mark.parametrize("object_kind", tuple(_OBJECTS))
def test_all_caption_kinds_bind_all_captionable_object_kinds(
    caption_kind: str,
    object_kind: str,
) -> None:
    source = f"{_CAPTIONS[caption_kind]}\n\n{_OBJECTS[object_kind]}\n"

    analysis = analyze_markdown_semantics_v3(source, input_id="matrix.md")

    assert not analysis.has_errors
    [target] = analysis.projection["targets"]
    assert target["kind"] == caption_kind
    object_range = target["object_range"]
    assert source[object_range["start"] : object_range["end"]].strip() == _OBJECTS[object_kind]

    plan = prepare_runtime_semantics_v3(source, input_id="matrix.md")
    [owner] = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source), plan)
    assert owner["_docwen_v3_caption_target"]["kind"] == caption_kind
    assert {
        "figure": "paragraph",
        "table": "table",
        "equation": "block_math",
        "code_block": "block_code",
    }[object_kind] == owner["type"]


@pytest.mark.parametrize("caption_first", [True, False])
@pytest.mark.parametrize("blank_lines", [0, 1])
@pytest.mark.parametrize("caption_kind", tuple(_CAPTIONS))
@pytest.mark.parametrize("object_kind", tuple(_OBJECTS))
def test_caption_binding_accepts_both_source_orders_with_zero_or_one_blank_line(
    caption_first: bool,
    blank_lines: int,
    caption_kind: str,
    object_kind: str,
) -> None:
    caption = _CAPTIONS[caption_kind]
    object_source = _OBJECTS[object_kind]
    separator = "\n" * (blank_lines + 1)
    source = f"{caption}{separator}{object_source}\n" if caption_first else f"{object_source}{separator}{caption}\n"

    analysis = analyze_markdown_semantics_v3(source, input_id="spacing.md")

    assert not analysis.has_errors
    assert [(item["kind"], item["title"]) for item in analysis.projection["targets"]] == [(caption_kind, "Composite")]
    plan = prepare_runtime_semantics_v3(source, input_id="spacing.md")
    owners = [
        node
        for node in apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source), plan)
        if "_docwen_v3_caption_target" in node
    ]
    assert len(owners) == 1
    assert owners[0]["_docwen_v3_caption_target"]["kind"] == caption_kind
    assert (
        owners[0]["type"]
        == {
            "figure": "paragraph",
            "table": "table",
            "equation": "block_math",
            "code_block": "block_code",
        }[object_kind]
    )


@pytest.mark.parametrize("caption_first", [True, False])
def test_two_blank_lines_break_caption_ownership(caption_first: bool) -> None:
    caption = "Figure: Composite"
    table = _OBJECTS["table"]
    source = f"{caption}\n\n\n{table}\n" if caption_first else f"{table}\n\n\n{caption}\n"

    analysis = analyze_markdown_semantics_v3(source, input_id="spacing.md")

    assert analysis.has_errors
    assert analysis.projection["targets"] == []
    assert [item["code"] for item in analysis.diagnostics] == ["docwen.markdown.caption.object_mismatch"]


@pytest.mark.parametrize(
    "source",
    [
        "![left](left.png)\n\nFigure: Ambiguous\n\n| right |\n|---|\n| 1 |\n",
        "Figure: First\n\n| shared |\n|---|\n| 1 |\n\nTable: Second\n",
    ],
)
def test_ambiguous_caption_object_graph_fails_closed(source: str) -> None:
    analysis = analyze_markdown_semantics_v3(source, input_id="ambiguous.md")

    assert analysis.has_errors
    assert analysis.projection["targets"] == []
    assert all(item["code"] == "docwen.markdown.caption.object_mismatch" for item in analysis.diagnostics)


def test_chain_does_not_use_global_matching_to_resolve_a_locally_ambiguous_caption() -> None:
    source = "Figure: First\n\n| first |\n|---|\n| 1 |\n\nTable: Second\n\n![second](second.png)\n"

    analysis = analyze_markdown_semantics_v3(source, input_id="ambiguous-chain.md")

    assert analysis.has_errors
    assert [(item["kind"], item["title"]) for item in analysis.projection["targets"]] == [("figure", "First")]
    assert [item["code"] for item in analysis.diagnostics] == ["docwen.markdown.caption.object_mismatch"]
    plan = prepare_runtime_semantics_v3(source, input_id="ambiguous-chain.md")
    assert plan.analysis.has_errors
    with pytest.raises(ValueError, match="invalid v3 analysis"):
        apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source), plan)


def test_next_line_caption_id_and_post_block_object_id_keep_distinct_owners() -> None:
    source = "| A |\n|---|\n| 1 |\n^object-id\nFigure: Composite\n^caption-id\n"

    analysis = analyze_markdown_semantics_v3(source, input_id="ids.md")
    assert not analysis.has_errors
    assert analysis.projection["targets"][0]["id"] == "caption-id"
    assert analysis.projection["anchors"][0]["id"] == "object-id"

    plan = prepare_runtime_semantics_v3(source, input_id="ids.md")
    [owner] = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source), plan)
    assert owner["type"] == "table"
    assert owner["_docwen_v3_caption_target"]["id"] == "caption-id"
    assert owner["_docwen_v3_ordinary_anchor"]["id"] == "object-id"


@pytest.mark.parametrize("declaration", ["Code: ^snippet", "Code:\n^snippet"])
def test_empty_code_caption_is_valid_with_explicit_id(declaration: str) -> None:
    source = f"{declaration}\n```text\nx\n```\n"

    analysis = analyze_markdown_semantics_v3(source, input_id="code.md")

    assert not analysis.has_errors
    [target] = analysis.projection["targets"]
    assert (target["kind"], target["title"], target["id"]) == ("code_block", "", "snippet")
