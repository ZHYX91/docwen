"""Source numbering and semantic caches consume the same actual headings."""

from __future__ import annotations

from collections.abc import Iterator
from copy import deepcopy
from typing import Any, Literal

import pytest

from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_core.text.heading_numbering import HeadingFormatter
from docwen_plugin_markdown.common_utils import _apply_numbering
from docwen_plugin_markdown.document_semantics_v3_fenced_source import recover_fenced_logical_body_v3
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.number_suite_direct_semantics import analyze_markdown_semantics_v3
from docwen_plugin_markdown.runtime_semantics_v3 import (
    RuntimeSemanticsV3Unsupported,
    apply_runtime_semantics_v3,
    prepare_runtime_semantics_v3,
)
from docwen_plugin_markdown.to_docx.converter import _direct_heading_number_provider

pytestmark = pytest.mark.contract

_SCHEME = {
    "level_1": {"format": "{1.arabic_half} "},
    "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "},
    "level_7": {"format": "{7.arabic_half} "},
    "level_8": {"format": "{8.arabic_half} "},
    "level_9": {"format": "{9.arabic_half} "},
}


def _walk(nodes: list[dict[str, Any]]) -> Iterator[dict[str, Any]]:
    for node in nodes:
        yield node
        yield from _walk(node.get("children", []))


def _text(node: dict[str, Any]) -> str:
    return "".join(str(child.get("raw", "")) for child in _walk(node.get("children", [])))


@pytest.mark.parametrize(
    ("prefix", "visible", "number"),
    [
        ("Setext\n===\n\n", ["1 Setext", "2 Real"], "2"),
        ("First\nsecond\n===\n\n", ["1 Firstsecond", "2 Real"], "2"),
        ("> Setext\n> ===\n\n", ["1 Setext", "2 Real"], "2"),
        ("- Setext\n  ===\n\n", ["1 Setext", "2 Real"], "2"),
        ("-\n  Setext\n  ===\n\n", ["1 Setext", "2 Real"], "2"),
        ("1.\n   Setext\n   ===\n\n", ["1 Setext", "2 Real"], "2"),
        ("-\tSetext\n\t===\n\n", ["1 Setext", "2 Real"], "2"),
        ("\ufeff- Setext\n  ===\n\n", ["1 Setext", "2 Real"], "2"),
        ("#\n\n", ["1", "2 Real"], "2"),
        ("# ###\n\n", ["1", "2 Real"], "2"),
        ("***\n===\n\n", ["1 Real"], "1"),
        ("intro\n2. # Fake\n\n", ["1 Real"], "1"),
    ],
)
def test_direct_number_cache_matches_the_visible_numbering_inventory(
    prefix: str, visible: list[str], number: str
) -> None:
    source = prefix + "# Real ^real\n\nSee @[[#^real]].\n"
    extensions = MarkdownExtensions(captions_references=True, extended_headings=True)
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="headings.md",
        extensions=extensions,
        consumer_profile="number_suite_direct",
        heading_number_provider=_direct_heading_number_provider(_SCHEME),
    )
    assert not plan.analysis.has_errors
    target = next(item for item in plan.analysis.projection["targets"] if item.get("id") == "real")
    assert target["number"] == number
    [reference] = plan.analysis.projection["references"]
    assert (reference["resolution_status"], reference["target_id"], reference["cached_number"]) == (
        "resolved",
        "real",
        number,
    )
    numbered = _apply_numbering(plan.shielded_source, HeadingFormatter(_SCHEME), extended_headings=True)
    ast = apply_runtime_semantics_v3(parse_markdown_text(numbered, extensions=extensions), plan)
    assert [_text(node).strip() for node in _walk(ast) if node["type"] == "heading"] == visible
    [rendered_reference] = [node for node in _walk(ast) if node["type"] == "semantic_cross_reference"]
    assert rendered_reference["cached_number"] == number


@pytest.mark.parametrize(
    ("prefix", "marks", "expected"),
    [
        ("Setext\n===\n\n", "#", "2"),
        ("Setext\n===\n\n", "##", "1.1"),
        ("Setext\n---\n\n", "##", "2"),
        ("#\n\n", "#", "2"),
        ("#######\n\n", "#######", "2"),
        ("########\n\n", "########", "2"),
        ("#########\n\n", "#########", "2"),
    ],
)
def test_default_direct_counter_includes_non_target_and_empty_headings(prefix: str, marks: str, expected: str) -> None:
    source = prefix + f"{marks} Real ^real\n\nSee @[[#^real]].\n"
    analysis = analyze_markdown_semantics_v3(
        source,
        input_id="headings.md",
        extensions=MarkdownExtensions(captions_references=True, extended_headings=True),
        consumer_profile="number_suite_direct",
    )
    assert not analysis.has_errors
    target = next(item for item in analysis.projection["targets"] if item.get("id") == "real")
    assert target["number"] == expected
    assert analysis.projection["references"][0]["cached_number"] == expected
    assert not any(item["title"] == "Setext" for item in analysis.projection["targets"])


@pytest.mark.parametrize("level", [7, 8, 9])
@pytest.mark.parametrize("container", ["> ", "- ", "1. ", "- > "])
@pytest.mark.parametrize("enabled", [False, True])
def test_extended_heading_parser_contract_applies_inside_each_container(
    level: int, container: str, enabled: bool
) -> None:
    source = f"{container}{'#' * level} Seven\n"
    ast = parse_markdown_text(source, extensions=MarkdownExtensions(extended_headings=enabled))
    headings = [node for node in _walk(ast) if node["type"] == "heading"]
    if enabled:
        assert [(node["attrs"]["level"], _text(node)) for node in headings] == [(level, "Seven")]
    else:
        assert headings == []
        assert "".join(str(node.get("raw", "")) for node in _walk(ast)) == f"{'#' * level} Seven"


@pytest.mark.parametrize("level", [7, 8, 9])
@pytest.mark.parametrize("container", ["> ", "- ", "1. ", "- > "])
@pytest.mark.parametrize("addressable", [False, True])
def test_extended_container_heading_plan_matches_the_actual_parser(
    level: int, container: str, addressable: bool
) -> None:
    source = f"{container}{'#' * level} Seven{' ^id' if addressable else ''}\n"
    extensions = MarkdownExtensions(extended_headings=True, captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="headings.md", extensions=extensions, consumer_profile="number_suite_direct"
    )
    assert not plan.analysis.has_errors
    [target] = plan.analysis.projection["targets"]
    assert (target["heading_level"], target["title"], target.get("id")) == (
        level,
        "Seven",
        "id" if addressable else None,
    )
    ast = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source, extensions=extensions), plan)
    [heading] = [node for node in _walk(ast) if node["type"] == "heading"]
    assert (heading["attrs"]["level"], _text(heading).strip()) == (level, "Seven")
    if addressable:
        assert heading["_docwen_v3_heading_target"]["id"] == "id"


def test_a_disabled_number_template_still_advances_the_shared_heading_hierarchy() -> None:
    source = "First\n===\n\nSecond\n===\n\n## Real ^real\n\nSee @[[#^real]].\n"
    scheme = {"level_1": {"format": ""}, "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "}}
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="headings.md",
        extensions=extensions,
        consumer_profile="number_suite_direct",
        heading_number_provider=_direct_heading_number_provider(scheme),
    )
    assert not plan.analysis.has_errors
    [target] = plan.analysis.projection["targets"]
    assert target["number"] == "2.1"
    numbered = _apply_numbering(plan.shielded_source, HeadingFormatter(scheme))
    ast = apply_runtime_semantics_v3(parse_markdown_text(numbered, extensions=extensions), plan)
    assert [_text(node).strip() for node in _walk(ast) if node["type"] == "heading"] == ["First", "Second", "2.1 Real"]
    assert plan.analysis.projection["references"][0]["cached_number"] == "2.1"


@pytest.mark.parametrize("marks", ["```", "`````", "~~~"])
@pytest.mark.parametrize("eol", ["\n", "\r\n"])
def test_direct_fenced_carrier_ends_at_the_actual_list_owner(marks: str, eol: str) -> None:
    source = (f"- text\nlazy\n  {marks}text\n  hidden\n\n# Real ^real\n\nSee @[[#^real]].\n").replace("\n", eol)
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="headings.md", extensions=extensions, consumer_profile="number_suite_direct"
    )

    assert not plan.analysis.has_errors
    [target] = plan.analysis.projection["targets"]
    assert (target["title"], target["id"], target["number"]) == ("Real", "real", "1")
    assert plan.analysis.projection["references"][0]["cached_number"] == "1"
    [carrier] = plan.analysis.projection["fenced_sources"]
    assert (carrier["source_start"], carrier["source_end"], carrier["closing_state"]) == (
        source.index(f"  {marks}text"),
        source.index("# Real"),
        "omitted_eof",
    )
    assert recover_fenced_logical_body_v3(source, carrier) == f"hidden{eol}{eol}"
    ast = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source, extensions=extensions), plan)
    [code] = [node for node in _walk(ast) if node["type"] == "block_code"]
    # Mistune trims the list separator; the authenticated source carrier above
    # independently retains the authored blank line for lossless recovery.
    assert code["raw"] == "hidden\n"
    assert code["_docwen_v3_fenced_source"] == carrier
    assert [_text(node).strip() for node in _walk(ast) if node["type"] == "heading"] == ["Real"]


@pytest.mark.parametrize("marks", ["```", "~~~~"])
@pytest.mark.parametrize("captions", [False, True])
def test_pseudo_fence_suffix_in_a_list_keeps_its_safe_rejection(marks: str, captions: bool) -> None:
    source = f"- text\nlazy\n  {marks}text\n  hidden\n  {marks} ^bad-code\n\n# Real ^real\n"
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="headings.md",
        extensions=MarkdownExtensions(captions_references=captions),
        consumer_profile="number_suite_direct" if captions else "frozen_v3",
    )

    assert plan.analysis.has_errors
    assert plan.analysis.projection["fenced_sources"] == []
    assert plan.markers == ()
    [diagnostic] = plan.analysis.diagnostics
    assert diagnostic["code"] == "docwen.markdown.anchor.invalid_id"
    assert source[diagnostic["range"]["start"] : diagnostic["range"]["end"]] == "^bad-code"


@pytest.mark.parametrize("start", [2, 3, 42])
def test_non_interrupting_ordered_spelling_does_not_manufacture_a_heading_target(start: int) -> None:
    source = f"intro\n{start}. # Fake ^ordinary\n\n# Real ^real\n\nSee @[[#^real]].\n"
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="headings.md", extensions=extensions, consumer_profile="number_suite_direct"
    )

    assert not plan.analysis.has_errors
    [target] = plan.analysis.projection["targets"]
    assert (target["title"], target["id"], target["number"]) == ("Real", "real", "1")
    [anchor] = plan.analysis.projection["anchors"]
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == ("ordinary", "paragraph", "inline")
    ast = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source, extensions=extensions), plan)
    assert [_text(node).strip() for node in _walk(ast) if node["type"] == "heading"] == ["Real"]
    assert _text(ast[0]).strip() == f"intro{start}. # Fake"
    [reference] = [node for node in _walk(ast) if node["type"] == "semantic_cross_reference"]
    assert reference["cached_number"] == "1"


def test_parser_ignores_only_the_initial_encoding_signature() -> None:
    source = "\ufeff# First\ufeffinside\r\n\r\nText\ufeffinside\r\n"
    ast = parse_markdown_text(source, extensions=MarkdownExtensions())
    assert [(node["type"], _text(node)) for node in ast if node["type"] != "blank_line"] == [
        ("heading", "First\ufeffinside"),
        ("paragraph", "Text\ufeffinside"),
    ]


@pytest.mark.parametrize("profile", ["frozen_v3", "number_suite_direct"])
@pytest.mark.parametrize("formula", ["$$ E = mc^2 $$\n", "$$\nE = mc^2\n$$\n"])
def test_canonical_opaque_inventory_preserves_actual_display_math_ownership(
    profile: Literal["frozen_v3", "number_suite_direct"], formula: str
) -> None:
    source = f"Equation: Energy ^energy\n\n{formula}^formula\n"
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="headings.md",
        extensions=MarkdownExtensions(captions_references=True),
        consumer_profile=profile,
    )

    assert not plan.analysis.has_errors
    [target] = plan.analysis.projection["targets"]
    assert (target["kind"], target["id"]) == ("equation", "energy")
    assert target["object_range"] == {"start": source.index("$$"), "end": source.index("^formula")}
    [anchor] = plan.analysis.projection["anchors"]
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == ("formula", "equation", "post_block")
    ast = apply_runtime_semantics_v3(
        parse_markdown_text(plan.shielded_source, extensions=MarkdownExtensions(captions_references=True)), plan
    )
    [math] = [node for node in _walk(ast) if node["type"] == "block_math"]
    assert math["raw"] == "E = mc^2"
    assert math["_docwen_v3_ordinary_anchor"]["id"] == "formula"


@pytest.mark.parametrize(("opening", "closing"), [("```text", "```"), ("<pre>", "</pre>")])
def test_math_examples_inside_opaque_blocks_do_not_acquire_equation_owners(opening: str, closing: str) -> None:
    source = f"{opening}\n$$\nEquation: Sample ^sample\n$$\n^formula\n{closing}\n\n# Real ^real\n"
    analysis = analyze_markdown_semantics_v3(
        source,
        input_id="headings.md",
        extensions=MarkdownExtensions(captions_references=True),
        consumer_profile="number_suite_direct",
    )
    assert not analysis.has_errors
    [target] = analysis.projection["targets"]
    assert (target["kind"], target["id"]) == ("heading", "real")
    assert analysis.projection["anchors"] == []


@pytest.mark.parametrize("prefix", ["$$\nUnclosed formula\n\n", "    $$\n    E = mc^2\n    $$\n\n"])
def test_math_like_text_without_a_block_owner_does_not_swallow_a_real_heading(prefix: str) -> None:
    source = prefix + "# Real ^real\n"
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="headings.md", extensions=extensions, consumer_profile="number_suite_direct"
    )
    assert not plan.analysis.has_errors
    [target] = plan.analysis.projection["targets"]
    assert (target["kind"], target["id"]) == ("heading", "real")
    ast = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source, extensions=extensions), plan)
    assert not any(node["type"] == "block_math" for node in _walk(ast))
    assert [_text(node).strip() for node in _walk(ast) if node["type"] == "heading"] == ["Real"]


@pytest.mark.parametrize("profile", ["frozen_v3", "number_suite_direct"])
@pytest.mark.parametrize("container", ["quote", "list"])
@pytest.mark.parametrize("multiline", [False, True])
def test_container_display_math_keeps_its_caption_anchor_and_literal_body(
    profile: Literal["frozen_v3", "number_suite_direct"], container: str, multiline: bool
) -> None:
    formula = ["$$", "E = mc^2", "# Literal ^sample", "$$"] if multiline else ["$$ E = mc^2 $$"]
    lines = (["Equation: Energy ^energy", ""] if profile == "frozen_v3" else []) + formula + ["^formula"]
    if container == "quote":
        source = "".join("> " + line + "\n" for line in lines)
    else:
        source = "- " + lines[0] + "\n" + "".join("  " + line + "\n" for line in lines[1:])
    source += "\n# Real ^real\n"
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(source, input_id="headings.md", extensions=extensions, consumer_profile=profile)

    assert not plan.analysis.has_errors
    assert [(target["kind"], target["id"]) for target in plan.analysis.projection["targets"]] == (
        [("equation", "energy"), ("heading", "real")] if profile == "frozen_v3" else [("heading", "real")]
    )
    [anchor] = plan.analysis.projection["anchors"]
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == ("formula", "equation", "post_block")
    ast = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source, extensions=extensions), plan)
    [math] = [node for node in _walk(ast) if node["type"] == "block_math"]
    assert math["raw"] == ("E = mc^2\n# Literal ^sample" if multiline else "E = mc^2")
    assert math["_docwen_v3_ordinary_anchor"]["id"] == "formula"
    assert [_text(node).strip() for node in _walk(ast) if node["type"] == "heading"] == ["Real"]


@pytest.mark.parametrize(
    "opening",
    [
        "- | Unique header |\n  | --- |\n  | Unique value |\n",
        "- ```text\n  literal body\n  ```\n",
        "- $$\n  x + 1 = y\n  $$\n",
        "- # Lead heading\n",
        "- Lead paragraph.\n",
    ],
    ids=["table", "fence", "math", "heading", "paragraph"],
)
def test_list_continuation_anchor_owns_only_its_own_paragraph(opening: str) -> None:
    continuation = "  Continuation after owner. ^continuation\n"
    source = opening + "\n" + continuation + "\nOutside after list.\n"
    extensions = MarkdownExtensions()
    plan = prepare_runtime_semantics_v3(source, input_id="list.md", extensions=extensions)

    assert not plan.analysis.has_errors
    [anchor] = plan.analysis.projection["anchors"]
    paragraph_start = source.index(continuation)
    paragraph_range = {"start": paragraph_start, "end": paragraph_start + len(continuation)}
    container_range = {"start": 0, "end": source.index("Outside after list.")}
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == ("continuation", "paragraph", "inline")
    assert anchor["block_range"] == paragraph_range
    assert anchor["range"] == {"start": source.index("^continuation"), "end": source.index("^continuation") + 13}
    assert anchor["container_path"] == [
        {"block_kind": "list", "block_range": container_range},
        {"block_kind": "list_item", "block_range": container_range},
    ]
    ast = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source, extensions=extensions), plan)
    [owner] = [node for node in _walk(ast) if "_docwen_v3_ordinary_anchor" in node]
    assert owner["type"] == "paragraph"
    assert _text(owner) == "Continuation after owner."
    assert owner["_docwen_v3_ordinary_anchor"] == anchor
    assert owner["_docwen_v3_ordinary_anchor_parent_source_id"] is None


@pytest.mark.parametrize("lead", ["- Lead ^item\n", "-\n  Lead ^item\n"])
def test_actual_list_lead_anchor_retains_the_full_item_and_parent_topology(lead: str) -> None:
    continuation = "  Continuation. ^child\n"
    source = lead + "\n  ```text\n  literal body\n  ```\n\n" + continuation + "\nOutside.\n"
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="list.md", extensions=extensions, consumer_profile="number_suite_direct"
    )

    assert not plan.analysis.has_errors
    item, child = plan.analysis.projection["anchors"]
    item_range = {"start": 0, "end": source.index("Outside.")}
    assert (item["id"], item["block_kind"], item["block_range"]) == ("item", "list_item", item_range)
    assert item["container_path"] == [{"block_kind": "list", "block_range": item_range}]
    child_start = source.index(continuation)
    assert (child["id"], child["block_kind"], child["block_range"]) == (
        "child",
        "paragraph",
        {"start": child_start, "end": child_start + len(continuation)},
    )
    assert child["container_path"] == [
        {"block_kind": "list", "block_range": item_range},
        {"block_kind": "list_item", "block_range": item_range},
    ]
    ast = apply_runtime_semantics_v3(parse_markdown_text(plan.shielded_source, extensions=extensions), plan)
    owners = {
        node["_docwen_v3_ordinary_anchor"]["id"]: node for node in _walk(ast) if "_docwen_v3_ordinary_anchor" in node
    }
    assert owners["item"]["type"] == "list_item"
    assert owners["item"]["_docwen_v3_ordinary_anchor_parent_source_id"] is None
    assert owners["child"]["type"] == "paragraph"
    assert _text(owners["child"]) == "Continuation."
    assert owners["child"]["_docwen_v3_ordinary_anchor_parent_source_id"] == "item"


def test_duplicated_list_lead_marker_is_rejected_without_choosing_an_owner() -> None:
    source = "- Lead ^item\n\n  Continuation. ^child\n"
    extensions = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="list.md", extensions=extensions, consumer_profile="number_suite_direct"
    )
    assert not plan.analysis.has_errors
    ast = parse_markdown_text(plan.shielded_source, extensions=extensions)
    [item] = ast[0]["children"]
    item["children"].insert(1, deepcopy(item["children"][0]))
    with pytest.raises(RuntimeSemanticsV3Unsupported, match="list item has multiple ordinary-anchor owners"):
        apply_runtime_semantics_v3(ast, plan)
