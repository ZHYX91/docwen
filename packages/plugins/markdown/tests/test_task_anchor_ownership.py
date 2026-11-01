"""Task-list source owners retain complete items and distinct child anchors."""

from __future__ import annotations

from collections.abc import Iterator
from copy import deepcopy
from typing import Any

import pytest

from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.runtime_semantics_v3 import (
    RuntimeSemanticsV3Unsupported,
    apply_runtime_semantics_v3,
    prepare_runtime_semantics_v3,
)

pytestmark = pytest.mark.contract


def _walk(nodes: list[dict[str, Any]]) -> Iterator[dict[str, Any]]:
    for node in nodes:
        yield node
        yield from _walk(node.get("children", []))


@pytest.mark.parametrize("lead", ["- Lead", "- [ ] Lead", "- [x] Lead"], ids=["ordinary", "unchecked", "checked"])
@pytest.mark.parametrize("direct", [False, True], ids=["default-dialect", "direct-dialect"])
def test_task_and_ordinary_item_owners_keep_later_child_identity(lead: str, direct: bool) -> None:
    source = lead + " ^item\n\n  Continuation. ^child\n\nOutside.\n"
    extensions = MarkdownExtensions(captions_references=direct)
    plan = prepare_runtime_semantics_v3(
        source,
        input_id="tasks.md",
        extensions=extensions,
        consumer_profile="number_suite_direct" if direct else "frozen_v3",
    )
    assert not plan.analysis.has_errors
    item, child = plan.analysis.projection["anchors"]
    item_range = {"start": 0, "end": source.index("Outside.")}
    continuation_start = source.index("  Continuation.")
    assert (item["id"], item["block_kind"], item["block_range"]) == ("item", "list_item", item_range)
    assert item["container_path"] == [{"block_kind": "list", "block_range": item_range}]
    assert (child["id"], child["block_kind"], child["block_range"]) == (
        "child",
        "paragraph",
        {"start": continuation_start, "end": source.index("\n", continuation_start) + 1},
    )
    assert child["container_path"] == [
        {"block_kind": "list", "block_range": item_range},
        {"block_kind": "list_item", "block_range": item_range},
    ]
    assert plan.ordinary_anchor_parents == (("item", None), ("child", "item"))
    parsed = parse_markdown_text(plan.shielded_source, extensions=extensions)
    assert isinstance(parsed, list)
    original_item = parsed[0]["children"][0]
    original_attrs = deepcopy(original_item.get("attrs", {}))
    ast = apply_runtime_semantics_v3(parsed, plan)
    owners = {
        node["_docwen_v3_ordinary_anchor"]["id"]: node for node in _walk(ast) if "_docwen_v3_ordinary_anchor" in node
    }
    assert set(owners) == {"item", "child"}
    expected_type = "list_item" if lead == "- Lead" else "task_list_item"
    assert owners["item"]["type"] == expected_type
    assert owners["item"].get("attrs", {}) == original_attrs
    if expected_type == "task_list_item":
        assert owners["item"]["attrs"]["checked"] is (lead == "- [x] Lead")
    assert owners["item"]["_docwen_v3_ordinary_anchor_parent_source_id"] is None
    assert owners["child"]["type"] == "paragraph"
    assert owners["child"]["_docwen_v3_ordinary_anchor_parent_source_id"] == "item"
    assert owners["child"] in list(_walk(owners["item"]["children"]))
    assert "".join(str(node.get("raw", "")) for node in _walk(owners["item"]["children"])) == "LeadContinuation."
    assert [node["type"] for node in owners["item"]["children"]] == ["paragraph", "blank_line", "paragraph"]


@pytest.mark.parametrize("lead", ["- Lead", "- [ ] Lead", "- [x] Lead"], ids=["ordinary", "unchecked", "checked"])
def test_duplicate_lead_owner_is_rejected_for_every_list_item_type(lead: str) -> None:
    source = lead + " ^item\n\n  Continuation. ^child\n\nOutside.\n"
    extensions = MarkdownExtensions()
    plan = prepare_runtime_semantics_v3(source, input_id="tasks.md", extensions=extensions)
    assert not plan.analysis.has_errors
    parsed = parse_markdown_text(plan.shielded_source, extensions=extensions)
    assert isinstance(parsed, list)
    [item] = parsed[0]["children"]
    item["children"].insert(1, deepcopy(item["children"][0]))
    with pytest.raises(RuntimeSemanticsV3Unsupported, match="list item has multiple ordinary-anchor owners"):
        apply_runtime_semantics_v3(parsed, plan)
