"""Source owner depth distinguishes equal-range ancestors from terminal self."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from typing import Any

import pytest

from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.runtime_semantics_v3 import (
    RuntimeSemanticsV3Unsupported,
    _ordinary_anchor_owner_path,
    _prepare_analyzed_runtime_plan,
    apply_runtime_semantics_v3,
    prepare_runtime_semantics_v3,
)

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("prefix", "kinds", "owner_kind"),
    [
        ("> > ", ("block_quote", "block_quote"), "paragraph"),
        ("> > > ", ("block_quote", "block_quote", "block_quote"), "paragraph"),
        ("> - ", ("block_quote", "list"), "list_item"),
        ("- > ", ("list", "list_item", "block_quote"), "paragraph"),
    ],
    ids=["double-quote", "triple-quote", "quote-list", "list-quote"],
)
@pytest.mark.parametrize(
    ("preamble", "newline"),
    [("", "\n"), ("\ufeff", "\r\n"), ("😀\r\n\r\n", "\r\n")],
    ids=["lf", "bom-crlf", "code-point-offset"],
)
def test_equal_range_source_ancestors_keep_every_structural_level(
    prefix: str, kinds: tuple[str, ...], owner_kind: str, preamble: str, newline: str
) -> None:
    source = preamble + prefix + "text ^id" + newline
    dialect = MarkdownExtensions()
    plan = prepare_runtime_semantics_v3(source, input_id="quote.md", extensions=dialect)
    assert not plan.analysis.has_errors
    [anchor] = plan.analysis.projection["anchors"]
    start = 0 if preamble in {"", "\ufeff"} else len(preamble)
    span = {"start": start, "end": len(source)}
    assert anchor["container_path"] == [{"block_kind": kind, "block_range": span} for kind in kinds]
    assert anchor["block_kind"] == owner_kind and anchor["block_range"] == span
    assert anchor["range"] == {"start": source.index("^id"), "end": source.index("^id") + 3}
    assert _ordinary_anchor_owner_path(anchor) == tuple((kind, start, len(source)) for kind in (*kinds, owner_kind))
    assert plan.ordinary_anchor_parents == (("id", None),)

    parsed = parse_markdown_text(plan.shielded_source, extensions=dialect, auto_link_bare_url=False)
    assert isinstance(parsed, list)
    restored = apply_runtime_semantics_v3(parsed, plan)
    owners: list[dict[str, Any]] = []

    def visit(nodes: list[dict[str, Any]]) -> None:
        for node in nodes:
            if "_docwen_v3_ordinary_anchor" in node:
                owners.append(node["_docwen_v3_ordinary_anchor"])
            visit(node.get("children", []))

    visit(restored)
    assert owners == [anchor]


def test_shorter_nested_quote_has_the_actual_source_range() -> None:
    source = "> a\n> > text ^id\n"
    plan = prepare_runtime_semantics_v3(source, input_id="shorter.md", extensions=MarkdownExtensions())
    [anchor] = plan.analysis.projection["anchors"]
    assert anchor["container_path"] == [
        {"block_kind": "block_quote", "block_range": {"start": 0, "end": len(source)}},
        {"block_kind": "block_quote", "block_range": {"start": len("> a\n"), "end": len(source)}},
    ]
    assert plan.ordinary_anchor_parents == (("id", None),)


def test_equal_range_ancestors_do_not_collapse_the_outer_anchor_parent() -> None:
    source = "> > text ^inner\n\n^outer\n"
    plan = prepare_runtime_semantics_v3(source, input_id="parent.md", extensions=MarkdownExtensions())
    inner, outer = plan.analysis.projection["anchors"]
    assert (inner["id"], outer["id"]) == ("inner", "outer")
    assert len(inner["container_path"]) == 2
    assert inner["container_path"][0] == inner["container_path"][1]
    assert inner["container_path"][0]["block_range"] == outer["block_range"]
    assert plan.ordinary_anchor_parents == (("inner", "outer"), ("outer", None))


def test_nested_quote_container_id_has_a_strictly_shorter_owner_than_its_ancestor() -> None:
    source = "> > text\n>\n> ^inner\n\n^outer\n"
    plan = prepare_runtime_semantics_v3(source, input_id="container.md", extensions=MarkdownExtensions())
    inner, outer = plan.analysis.projection["anchors"]
    assert (inner["id"], inner["block_kind"], inner["placement"]) == ("inner", "block_quote", "post_block")
    assert inner["block_range"] == {"start": 0, "end": len("> > text\n")}
    assert outer["block_range"] == {"start": 0, "end": source.index("\n^outer")}
    assert inner["container_path"] == [{"block_kind": "block_quote", "block_range": outer["block_range"]}]
    assert plan.ordinary_anchor_parents == (("inner", "outer"), ("outer", None))


@pytest.mark.parametrize(
    ("damage", "message"),
    [
        ("terminal-self", "contains its own block"),
        ("non-container-kind", "non-container kind"),
        ("crossing", "not laminar"),
        ("negative-start", "invalid range"),
        ("empty-range", "invalid range"),
        ("ambiguous-parent", "ambiguous direct parents"),
    ],
)
def test_invalid_owner_topology_is_rejected_before_marker_planning(damage: str, message: str) -> None:
    source = "> > text\n>\n> ^inner\n\n^outer\n"
    plan = prepare_runtime_semantics_v3(source, input_id="invalid.md", extensions=MarkdownExtensions())
    projection = deepcopy(plan.analysis.projection)
    inner, outer = projection["anchors"]
    if damage == "terminal-self":
        inner["container_path"].append({"block_kind": inner["block_kind"], "block_range": dict(inner["block_range"])})
    elif damage == "non-container-kind":
        inner["container_path"][0]["block_kind"] = "paragraph"
    elif damage == "crossing":
        inner["container_path"][0]["block_range"]["start"] = 1
    elif damage == "negative-start":
        inner["container_path"][0]["block_range"]["start"] = -1
    elif damage == "empty-range":
        inner["container_path"][0]["block_range"]["end"] = 0
    else:
        duplicate = deepcopy(outer)
        duplicate["id"] = "another-parent"
        projection["anchors"].append(duplicate)
    with pytest.raises(RuntimeSemanticsV3Unsupported, match=message):
        _prepare_analyzed_runtime_plan(source, replace(plan.analysis, projection=projection))
