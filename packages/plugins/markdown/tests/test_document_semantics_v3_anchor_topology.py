"""Source-oracle gates for nested ordinary-anchor owner paths."""

from __future__ import annotations

import pytest

from docwen_plugin_markdown.document_semantics_v3 import analyze_markdown_semantics_v3

pytestmark = pytest.mark.contract


def _anchors(source: str) -> list[dict]:
    analysis = analyze_markdown_semantics_v3(source, input_id="anchor-topology.md")
    assert not analysis.has_errors
    return analysis.projection["anchors"]


def test_multi_paragraph_list_item_owns_its_complete_structural_range() -> None:
    source = "- first ^inner-item\n\n  ```rust\n  body\n  ```\n\n- second\n\n^outer-list\n"

    inner, outer = _anchors(source)

    assert inner["container_path"] == [
        {
            "block_kind": "list",
            "block_range": outer["block_range"],
        }
    ]
    assert source[inner["block_range"]["start"] : inner["block_range"]["end"]] == (
        "- first ^inner-item\n\n  ```rust\n  body\n  ```\n\n"
    )
    assert outer["container_path"] == []


def test_equal_physical_quote_and_fence_owners_keep_distinct_source_paths() -> None:
    source = "> ```mermaid\n> graph TD\n> ```\n>\n> ^inner-fence\n\n^outer-quote\n"

    inner, outer = _anchors(source)

    assert (inner["id"], inner["block_kind"]) == ("inner-fence", "fenced_block")
    assert inner["container_path"] == [
        {
            "block_kind": "block_quote",
            "block_range": outer["block_range"],
        }
    ]
    assert (outer["id"], outer["block_kind"], outer["container_path"]) == (
        "outer-quote",
        "block_quote",
        [],
    )


@pytest.mark.parametrize(
    ("source", "inner_id", "inner_kind", "parent_id", "parent_kind"),
    [
        (
            "> - one\n> - two\n>\n> ^inner-list\n\n^outer-quote\n",
            "inner-list",
            "list",
            "outer-quote",
            "block_quote",
        ),
        (
            "- > quote\n  > continuation\n\n  ^inner-quote\n\n^outer-list\n",
            "inner-quote",
            "block_quote",
            "outer-list",
            "list",
        ),
    ],
)
def test_list_quote_nesting_projects_closed_outer_to_inner_segments(
    source: str,
    inner_id: str,
    inner_kind: str,
    parent_id: str,
    parent_kind: str,
) -> None:
    inner, parent = _anchors(source)

    assert (inner["id"], inner["block_kind"]) == (inner_id, inner_kind)
    assert (parent["id"], parent["block_kind"], parent["container_path"]) == (
        parent_id,
        parent_kind,
        [],
    )
    assert inner["container_path"][0] == {
        "block_kind": parent_kind,
        "block_range": parent["block_range"],
    }
    assert all(set(segment) == {"block_kind", "block_range"} for segment in inner["container_path"])


def test_disjoint_top_level_paragraph_anchors_have_no_container_or_parent_hint() -> None:
    first, second = _anchors("First ^first\n\nSecond ^second\n")

    assert first["container_path"] == []
    assert second["container_path"] == []
    assert first["block_range"]["end"] < second["block_range"]["start"]
    assert all("parent" not in key for item in (first, second) for key in item)
