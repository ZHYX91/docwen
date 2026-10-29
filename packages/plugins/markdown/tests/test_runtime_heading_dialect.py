"""Request-owned heading grammar must precede source-runtime shielding."""

from __future__ import annotations

import hashlib
import json

import pytest

from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_plugin_markdown.document_semantics_v3 import analyze_markdown_semantics_v3
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.runtime_semantics_v3 import (
    SemanticConsumerProfile,
    apply_runtime_semantics_v3,
    prepare_runtime_semantics_v3,
)

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("profile", ["frozen_v3", "number_suite_direct"])
@pytest.mark.parametrize("level", [7, 8, 9])
def test_disabled_extended_heading_preserves_the_complete_code_owner(
    profile: SemanticConsumerProfile, level: int
) -> None:
    literal = f"start\n{'#' * level} １２３ [link](https://example.com) @[[#^missing]] ^sample\nend"
    source = f"`{literal}`\n"
    extensions = MarkdownExtensions(captions_references=profile == "number_suite_direct")
    plan = prepare_runtime_semantics_v3(source, input_id="authored", extensions=extensions, consumer_profile=profile)

    assert not plan.analysis.has_errors
    assert plan.source_sha256 == hashlib.sha256(source.encode()).hexdigest()
    assert plan.analysis.projection["targets"] == []
    assert plan.analysis.projection["anchors"] == []
    assert plan.analysis.projection["references"] == []
    assert plan.markers == ()
    assert plan.shielded_source == source
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    assert ast == [{"type": "paragraph", "children": [{"type": "codespan", "raw": literal.replace("\n", " ")}]}]


@pytest.mark.parametrize("profile", ["frozen_v3", "number_suite_direct"])
@pytest.mark.parametrize("level", [7, 8, 9])
def test_enabled_extended_heading_owns_its_original_target(profile: SemanticConsumerProfile, level: int) -> None:
    source = f"`start\n{'#' * level} １２３ ^real\nend`\n"
    extensions = MarkdownExtensions(captions_references=profile == "number_suite_direct", extended_headings=True)
    plan = prepare_runtime_semantics_v3(source, input_id="authored", extensions=extensions, consumer_profile=profile)

    assert not plan.analysis.has_errors
    [target] = plan.analysis.projection["targets"]
    assert (target["kind"], target["heading_level"], target["id"], target["title"]) == (
        "heading",
        level,
        "real",
        "１２３",
    )
    assert source[target["id_range"]["start"] : target["id_range"]["end"]] == "^real"
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    assert [node["type"] for node in ast] == ["paragraph", "heading", "paragraph"]
    assert "".join(child.get("raw", "") for child in ast[0]["children"]) == "`start"
    assert "".join(child.get("raw", "") for child in ast[1]["children"]).strip() == "１２３"
    assert "".join(child.get("raw", "") for child in ast[2]["children"]) == "end`"
    assert ast[1]["_docwen_v3_heading_target"]["id"] == "real"
    assert "DOCWENV3X" not in json.dumps(ast)


def test_source_native_disabled_captions_keep_literal_policy_and_v3_schema() -> None:
    source = "Figure: Literal title ^literal\n\nSee @[[#^literal]].\n"
    extensions = MarkdownExtensions()
    plan = prepare_runtime_semantics_v3(source, input_id="authored", extensions=extensions)

    assert not plan.analysis.has_errors
    assert plan.analysis.projection["schema"] == "docwen.markdown_semantics.v3"
    assert plan.analysis.projection["targets"] == []
    assert plan.analysis.projection["references"] == []
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    paragraphs = [node for node in ast if node["type"] == "paragraph"]
    assert ["".join(child.get("raw", "") for child in node["children"]) for node in paragraphs] == [
        "Figure: Literal title ^literal",
        "See @[[#^literal]].",
    ]
    assert "DOCWENV3X" not in json.dumps(ast)


def test_unspecified_runtime_dialect_retains_the_frozen_oracle_projection() -> None:
    source = "`start\n####### Frozen heading ^real\nend`\n"
    oracle = analyze_markdown_semantics_v3(source, input_id="frozen")
    plan = prepare_runtime_semantics_v3(source, input_id="frozen")

    assert plan.analysis == oracle
    assert plan.analysis.projection["schema"] == "docwen.markdown_semantics.v3"
    [target] = plan.analysis.projection["targets"]
    assert (target["heading_level"], target["id"]) == (7, "real")


@pytest.mark.parametrize(("ticks", "keyword"), [("`", "Figure"), ("``", "Table")])
@pytest.mark.parametrize("captions", [False, True])
def test_caption_candidate_cannot_split_a_complete_code_owner(ticks: str, keyword: str, captions: bool) -> None:
    literal = f"start\n{keyword}: Literal title\nend"
    source = f"{ticks}{literal}{ticks}\n"
    extensions = MarkdownExtensions(captions_references=captions)
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="authored",
        extensions=extensions,
        consumer_profile="number_suite_direct" if captions else "frozen_v3",
    )

    assert not plan.analysis.has_errors
    assert plan.source_sha256 == hashlib.sha256(source.encode()).hexdigest()
    for key in ("targets", "anchors", "references", "citations", "fenced_sources"):
        assert plan.analysis.projection[key] == []
    assert plan.analysis.literal_ranges == ()
    assert plan.markers == ()
    assert plan.ordinary_anchor_parents == ()
    assert plan.shielded_source == source
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    assert ast == [{"type": "paragraph", "children": [{"type": "codespan", "raw": literal.replace("\n", " ")}]}]


@pytest.mark.parametrize(("ticks", "keyword"), [("`", "Figure"), ("``", "Table")])
def test_caption_after_code_keeps_its_own_source_range_and_counter(ticks: str, keyword: str) -> None:
    literal = f"start\n{keyword}: Sample ^hidden\nend"
    source = f"{ticks}{literal}{ticks}\n\n{keyword}: Actual ^real\n\nSee @[[#^real]].\n"
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="authored", extensions=extensions, consumer_profile="number_suite_direct"
    )

    assert not plan.analysis.has_errors
    [target] = plan.analysis.projection["targets"]
    assert (target["kind"], target["title"], target["id"], target["number"]) == (
        keyword.lower(),
        "Actual",
        "real",
        "1",
    )
    assert target["id_range"] == {"start": source.index("^real"), "end": source.index("^real") + 5}
    assert source[target["declaration_range"]["start"] : target["declaration_range"]["end"]] == (
        f"{keyword}: Actual ^real\n"
    )
    [reference] = plan.analysis.projection["references"]
    assert (reference["target_id"], reference["cached_number"], reference["resolution_status"]) == (
        "real",
        "1",
        "resolved",
    )
    assert [marker.role for marker in plan.markers] == ["caption_declaration", "cross_reference"]
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    assert ast[0] == {"type": "paragraph", "children": [{"type": "codespan", "raw": literal.replace("\n", " ")}]}
    assert "DOCWENV3X" not in json.dumps(ast)


@pytest.mark.parametrize("ticks", ["`", "``"])
@pytest.mark.parametrize("captions", [False, True])
def test_code_owned_anchor_candidates_do_not_hide_a_later_real_anchor(ticks: str, captions: bool) -> None:
    literal = "start\n^sample\nline ^inline\nend"
    source = f"{ticks}{literal}{ticks}\n\nActual paragraph ^real\n"
    extensions = MarkdownExtensions(captions_references=captions)
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="authored",
        extensions=extensions,
        consumer_profile="number_suite_direct" if captions else "frozen_v3",
    )

    assert not plan.analysis.has_errors
    assert plan.analysis.projection["targets"] == []
    [anchor] = plan.analysis.projection["anchors"]
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == ("real", "paragraph", "inline")
    assert anchor["range"] == {"start": source.index("^real"), "end": source.index("^real") + 5}
    assert [marker.role for marker in plan.markers] == ["ordinary_anchor"]
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    paragraphs = [node for node in ast if node["type"] == "paragraph"]
    assert paragraphs[0] == {
        "type": "paragraph",
        "children": [{"type": "codespan", "raw": literal.replace("\n", " ")}],
    }
    assert "".join(child.get("raw", "") for child in paragraphs[1]["children"]) == "Actual paragraph"
    assert paragraphs[1]["_docwen_v3_ordinary_anchor"]["id"] == "real"
    assert "DOCWENV3X" not in json.dumps(ast)


@pytest.mark.parametrize(
    ("prefix", "kind", "title"),
    [
        ("Text ", None, None),
        ("# Real ", "heading", "Real `literal ^sample`"),
        ("Figure: ", "figure", "`literal ^sample`"),
    ],
)
def test_inline_code_suffix_cannot_become_an_anchor(prefix: str, kind: str | None, title: str | None) -> None:
    source = f"{prefix}`literal ^sample`\n"
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="authored", extensions=extensions, consumer_profile="number_suite_direct"
    )

    assert not plan.analysis.has_errors
    assert plan.analysis.projection["anchors"] == []
    if kind is None:
        assert plan.analysis.projection["targets"] == []
    else:
        [target] = plan.analysis.projection["targets"]
        assert (target["kind"], target["title"]) == (kind, title)
        assert "id" not in target
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    assert ast[0]["children"][-1] == {"type": "codespan", "raw": "literal ^sample"}
    assert "DOCWENV3X" not in json.dumps(ast)


@pytest.mark.parametrize(
    ("source", "literal"),
    [
        ("> `start\n> Figure: Literal title\n> end`\n", "start Figure: Literal title end"),
        ("- ``start\n  Table: Literal title\n  end``\n", "start Table: Literal title end"),
    ],
)
def test_nested_caption_candidates_keep_the_complete_code_owner(source: str, literal: str) -> None:
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="authored", extensions=extensions, consumer_profile="number_suite_direct"
    )

    assert not plan.analysis.has_errors
    assert plan.analysis.projection["targets"] == []
    assert plan.analysis.projection["anchors"] == []
    assert plan.markers == ()
    assert plan.shielded_source == source
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    pending = list(ast)
    codespans = []
    while pending:
        node = pending.pop()
        if node["type"] == "codespan":
            codespans.append(node["raw"])
        pending.extend(node.get("children", []))
    assert codespans == [literal]
    assert "DOCWENV3X" not in json.dumps(ast)


@pytest.mark.parametrize("captions", [False, True])
def test_source_literal_ownership_keeps_the_real_fence_carrier(captions: bool) -> None:
    body = "Figure: literal\n^sample\n"
    source = f"```text\n{body}```\n^code\n"
    extensions = MarkdownExtensions(captions_references=captions)
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="authored",
        extensions=extensions,
        consumer_profile="number_suite_direct" if captions else "frozen_v3",
    )

    assert not plan.analysis.has_errors
    assert plan.analysis.projection["targets"] == []
    [anchor] = plan.analysis.projection["anchors"]
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == ("code", "code_block", "post_block")
    assert anchor["range"] == {"start": source.index("^code"), "end": source.index("^code") + 5}
    [record] = plan.analysis.projection["fenced_sources"]
    assert (record["source_start"], record["source_end"], record["closing_state"]) == (
        0,
        source.index("^code"),
        "present",
    )
    assert [marker.role for marker in plan.markers] == ["fenced_source", "ordinary_anchor"]
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, auto_link_bare_url=False, extensions=extensions), plan
    )
    [code] = ast
    assert code["type"] == "block_code"
    assert code["raw"] == body
    assert code["_docwen_v3_fenced_source"] == record
    assert code["_docwen_v3_ordinary_anchor"]["id"] == "code"


@pytest.mark.parametrize("captions", [False, True])
def test_source_literal_ownership_keeps_invalid_fence_suffix_rejection(captions: bool) -> None:
    source = "```text\nbody\n``` ^bad-code\n\nafter\n"
    extensions = MarkdownExtensions(captions_references=captions)
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="authored",
        extensions=extensions,
        consumer_profile="number_suite_direct" if captions else "frozen_v3",
    )

    assert plan.analysis.has_errors
    assert plan.analysis.projection["fenced_sources"] == []
    assert plan.markers == ()
    [diagnostic] = plan.analysis.diagnostics
    assert diagnostic["code"] == "docwen.markdown.anchor.invalid_id"
    assert source[diagnostic["range"]["start"] : diagnostic["range"]["end"]] == "^bad-code"
