"""Focused production-adapter tests for the Markdown semantics v3 slice."""

from __future__ import annotations

import hashlib

import pytest

from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.runtime_semantics_v3 import (
    RuntimeSemanticsV3Unsupported,
    apply_runtime_semantics_v3,
    prepare_runtime_semantics_v3,
)

pytestmark = pytest.mark.contract


def test_runtime_adapter_preserves_five_observable_source_constructs() -> None:
    source = """# Intro ^intro

# Other

Plain block ^raw

See [[Page#^raw]] and ![[Page#^raw]], @[[#^intro]], @[[#Other|Current title]], and @fig-legacy.
"""

    plan = prepare_runtime_semantics_v3(source, input_id="source")
    assert not plan.analysis.has_errors
    assert plan.shielded_source != source
    assert all(item.marker not in source for item in plan.markers)

    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )
    headings = [item for item in ast if item["type"] == "heading"]
    assert headings[0]["_docwen_v3_heading_target"]["id"] == "intro"
    assert "_docwen_v3_heading_target" not in headings[1]
    anchored = next(item for item in ast if item.get("_docwen_v3_ordinary_anchor"))
    assert anchored["_docwen_v3_ordinary_anchor"] == {
        "id": "raw",
        "block_kind": "paragraph",
        "placement": "inline",
        "range": {"start": 37, "end": 41},
        "block_range": {"start": 25, "end": 42},
        "container_path": [],
    }

    paragraph = ast[-1]
    # Ordinary WikiLinks remain visible to the request-scoped link policy;
    # only semantic references and citations are shielded from preprocessing.
    inline_text = "".join(str(item.get("raw", "")) for item in paragraph["children"] if item["type"] == "text")
    assert "[[Page#^raw]]" in inline_text
    assert "![[Page#^raw]]" in inline_text
    references = [item for item in paragraph["children"] if item["type"] == "semantic_cross_reference"]
    assert [(item["selector_kind"], item["cached_number"], item.get("alias")) for item in references] == [
        ("stable_id", "1", None),
        ("heading_path", "2", "Current title"),
    ]
    citations = [item for item in paragraph["children"] if item["type"] == "semantic_citation"]
    assert citations[0]["raw"] == "@fig-legacy"
    assert citations[0]["items"][0]["key"] == "fig-legacy"


def test_runtime_adapter_binds_figure_caption_and_whole_list_anchor() -> None:
    source = "Figure: Caption ^figure\n\n![image](pixel.png)\n\n- one\n- two\n\n^whole-list\n"
    plan = prepare_runtime_semantics_v3(source, input_id="source")
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )

    figure = next(item for item in ast if item.get("_docwen_v3_caption_target"))
    assert figure["_docwen_v3_caption_target"]["kind"] == "figure"
    assert figure["_docwen_v3_caption_target"]["id"] == "figure"
    whole_list = next(item for item in ast if item.get("_docwen_v3_ordinary_anchor"))
    assert whole_list["type"] == "list"
    assert whole_list["_docwen_v3_ordinary_anchor"]["id"] == "whole-list"
    assert whole_list["_docwen_v3_ordinary_anchor"]["block_kind"] == "list"


def test_runtime_adapter_lifts_inline_anchor_to_its_list_item() -> None:
    source = "- first ^first-item\n- second\n"
    plan = prepare_runtime_semantics_v3(source, input_id="source")

    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )

    first, second = ast[0]["children"]
    assert first["_docwen_v3_ordinary_anchor"]["id"] == "first-item"
    assert "_docwen_v3_ordinary_anchor" not in second


def test_caption_title_markers_do_not_overlap_and_image_keeps_ordinary_anchor() -> None:
    source = "Figure: Caption @cite and [[Page#Heading]] ^semantic\n\n![x](a.png) ^raw\n"
    plan = prepare_runtime_semantics_v3(source, input_id="source")

    assert not plan.analysis.has_errors
    assert "[[Page#Heading]]" in plan.shielded_source
    assert "\n\n![x](a.png)" in plan.shielded_source
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )

    [figure] = ast
    assert figure["_docwen_v3_caption_target"]["id"] == "semantic"
    assert figure["_docwen_v3_caption_target"]["title"] == "Caption @cite and [[Page#Heading]]"
    assert figure["_docwen_v3_ordinary_anchor"]["id"] == "raw"
    assert figure["_docwen_v3_ordinary_anchor"]["block_kind"] == "image"


def test_nested_container_binding_is_recursive_and_transparent() -> None:
    source = """> # Nested ^heading
>
> Figure: Nested @cite ^figure
>
> ![x](a.png) ^raw-image
>
> ```mermaid
> graph TD
> ```
>
> ^raw-graph
"""
    plan = prepare_runtime_semantics_v3(source, input_id="source")

    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )

    [quote] = ast
    heading = next(item for item in quote["children"] if item["type"] == "heading")
    figure = next(item for item in quote["children"] if item.get("_docwen_v3_caption_target"))
    fenced = next(item for item in quote["children"] if item["type"] == "block_code")
    assert heading["_docwen_v3_heading_target"]["id"] == "heading"
    assert figure["_docwen_v3_caption_target"]["id"] == "figure"
    assert figure["_docwen_v3_ordinary_anchor"]["id"] == "raw-image"
    assert fenced["_docwen_v3_ordinary_anchor"]["id"] == "raw-graph"
    assert fenced["_docwen_v3_ordinary_anchor"]["block_kind"] == "fenced_block"


def test_runtime_derives_equal_range_anchor_parent_from_source_owner_paths() -> None:
    source = "> ```mermaid\n> graph TD\n> ```\n>\n> ^inner-fence\n\n^outer-quote\n"
    plan = prepare_runtime_semantics_v3(source, input_id="equal-range.md")

    assert plan.ordinary_anchor_parents == (
        ("inner-fence", "outer-quote"),
        ("outer-quote", None),
    )
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )
    [quote] = [item for item in ast if item["type"] == "block_quote"]
    [fenced] = [item for item in quote["children"] if item["type"] == "block_code"]
    assert fenced["_docwen_v3_ordinary_anchor_parent_source_id"] == "outer-quote"
    assert quote["_docwen_v3_ordinary_anchor_parent_source_id"] is None


def test_disjoint_top_level_anchors_have_no_runtime_topology_edges() -> None:
    source = "First ^first\n\nSecond ^second\n"
    plan = prepare_runtime_semantics_v3(source, input_id="disjoint.md")

    assert plan.ordinary_anchor_parents == (("first", None), ("second", None))
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )
    owners = [item for item in ast if item.get("_docwen_v3_ordinary_anchor")]
    assert [item["_docwen_v3_ordinary_anchor_parent_source_id"] for item in owners] == [None, None]


def test_single_list_item_nested_heading_is_bound_recursively() -> None:
    source = "- # Nested heading ^nested-heading\n"
    plan = prepare_runtime_semantics_v3(source, input_id="source")

    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )

    heading = ast[0]["children"][0]["children"][0]
    assert heading["type"] == "heading"
    assert heading["_docwen_v3_heading_target"]["id"] == "nested-heading"


def test_nested_paragraph_anchor_stays_on_the_paragraph_not_the_quote() -> None:
    source = "> paragraph ^inside\n"
    plan = prepare_runtime_semantics_v3(source, input_id="source")

    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )

    [quote] = ast
    [paragraph] = quote["children"]
    assert "_docwen_v3_ordinary_anchor" not in quote
    assert paragraph["_docwen_v3_ordinary_anchor"]["id"] == "inside"
    assert paragraph["_docwen_v3_ordinary_anchor"]["block_kind"] == "paragraph"


def test_runtime_plan_masks_yaml_from_semantics_but_keeps_full_source_identity() -> None:
    source = '---\ntitle: "@yaml [[Page#Heading]] ^bad_id"\n---\n\n# 正文😀 ^heading\n'
    plan = prepare_runtime_semantics_v3(source, input_id="machine-source")

    assert not plan.analysis.has_errors
    assert plan.source_sha256 == hashlib.sha256(source.encode("utf-8")).hexdigest()
    assert plan.shielded_source[: plan.body_start] == source[: plan.body_start]
    assert plan.shielded_body.startswith("\n# 正文😀 ")
    assert "@yaml" not in "".join(item["raw"] for item in plan.analysis.projection["citations"])
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_body, auto_link_bare_url=False),
        plan,
    )
    heading = next(item for item in ast if item["type"] == "heading")
    assert heading["_docwen_v3_heading_target"]["id"] == "heading"


def test_runtime_adapter_fails_closed_without_external_neutral_resolver() -> None:
    with pytest.raises(RuntimeSemanticsV3Unsupported, match="external neutral resolver"):
        prepare_runtime_semantics_v3("# Intro\n\n@[[Other#Intro]]\n", input_id="source")


def test_runtime_adapter_wraps_source_preflight_rejection_with_cause() -> None:
    with pytest.raises(RuntimeSemanticsV3Unsupported, match="only LF and CRLF") as raised:
        prepare_runtime_semantics_v3("```rust\rbody\r```\r", input_id="unsupported-eol.md")

    assert isinstance(raised.value.__cause__, ValueError)


def test_invalid_source_returns_exact_oracle_diagnostic_before_ast_rewrite() -> None:
    plan = prepare_runtime_semantics_v3("Plain ^bad_id\n", input_id="source")

    assert plan.analysis.has_errors
    assert plan.markers == ()
    assert plan.shielded_source == "Plain ^bad_id\n"
    [diagnostic] = plan.analysis.diagnostics
    assert diagnostic["code"] == "docwen.markdown.anchor.invalid_id"
    assert diagnostic["range"] == {"start": 6, "end": 13}


def test_truncated_invalid_fence_diagnostic_never_becomes_runtime_carrier_evidence() -> None:
    source = "```text\nbody\n``` ^bad-code\n\nafter\n"

    plan = prepare_runtime_semantics_v3(source, input_id="invalid-fence-suffix.md")

    assert plan.analysis.has_errors
    assert plan.analysis.projection["fenced_sources"] == []
    assert plan.markers == ()
    [diagnostic] = plan.analysis.diagnostics
    assert diagnostic["code"] == "docwen.markdown.anchor.invalid_id"
    assert source[diagnostic["range"]["start"] : diagnostic["range"]["end"]] == "^bad-code"


def test_commonmark_invalid_backtick_fence_opener_does_not_leak_core_errors() -> None:
    source = "```rust`invalid\nbody\n```\n"

    plan = prepare_runtime_semantics_v3(source, input_id="invalid-opener.md")
    [record] = plan.analysis.projection["fenced_sources"]
    assert record["source_start"] == source.rindex("```")
    assert record["closing_state"] == "omitted_eof"

    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False),
        plan,
    )
    assert [node["type"] for node in ast] == ["paragraph", "block_code"]
    assert ast[1]["_docwen_v3_fenced_source"] == record


def test_marker_namespace_is_source_collision_safe_and_deterministic() -> None:
    source = "# Intro ^intro\n\n@[[#^intro]]\n"
    first = prepare_runtime_semantics_v3(source, input_id="source")
    second = prepare_runtime_semantics_v3(source, input_id="source")

    assert first.shielded_source == second.shielded_source
    assert first.markers == second.markers
    assert all(item.marker.isascii() and item.marker.isalnum() for item in first.markers)
