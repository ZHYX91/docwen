from __future__ import annotations

import hashlib
from typing import Any

import pytest

from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.resolved_source_carriers_v4 import (
    apply_resolved_source_carriers_v4,
    prepare_resolved_source_carriers_v4,
)
from docwen_plugin_markdown.runtime_semantics_v3 import RuntimeSemanticsV3Unsupported

pytestmark = pytest.mark.unit


def _sha256(value: str) -> str:
    return hashlib.sha256(value.encode()).hexdigest()


def _walk(nodes: list[dict[str, Any]]):
    for node in nodes:
        yield node
        children = node.get("children")
        if isinstance(children, list):
            yield from _walk(children)


def test_carrier_bridge_keeps_numbering_and_resolution_profile_free() -> None:
    source = (
        "## 2.3 标题 ^head\n\n"
        "@[[#^head]] and [[Page]] and [@smith]\n\n"
        "> ```mermaid\n"
        "> graph TD\n"
        "> ```\n"
        ">\n"
        "> ^inner-fence\n\n"
        "^outer-quote\n"
    )

    plan = prepare_resolved_source_carriers_v4(
        source,
        input_id="probe.md",
        expected_source_sha256=_sha256(source),
    )

    assert {marker.role for marker in plan.runtime_plan.markers} == {"ordinary_anchor", "fenced_source"}
    assert "2.3 标题 ^head" in plan.shielded_source
    assert "@[[#^head]]" in plan.shielded_source
    assert "[[Page]]" in plan.shielded_source
    assert "[@smith]" in plan.shielded_source

    ast = parse_markdown_text(plan.shielded_source, auto_link_bare_url=False)
    restored = apply_resolved_source_carriers_v4(ast, plan)
    nodes = tuple(_walk(restored))
    heading = next(item for item in nodes if item.get("type") == "heading")
    assert heading["children"] == [{"type": "text", "raw": "2.3 标题 ^head"}]
    fence = next(item for item in nodes if item.get("type") == "block_code")
    assert fence["attrs"] == {"info": "mermaid"}
    assert fence["_docwen_v3_fenced_body"] == "graph TD\n"
    assert fence["_docwen_v3_ordinary_anchor"]["id"] == "inner-fence"
    quote = next(item for item in nodes if item.get("type") == "block_quote")
    assert quote["_docwen_v3_ordinary_anchor"]["id"] == "outer-quote"
    assert fence["_docwen_v3_ordinary_anchor_parent_source_id"] == "outer-quote"


@pytest.mark.parametrize(
    "mutated_hash",
    ["0" * 64, "f" * 64],
)
def test_carrier_bridge_rejects_a_different_source_identity(mutated_hash: str) -> None:
    source = "Text ^raw\n"
    with pytest.raises(RuntimeSemanticsV3Unsupported, match="different authenticated source"):
        prepare_resolved_source_carriers_v4(
            source,
            input_id="probe.md",
            expected_source_sha256=mutated_hash,
        )


@pytest.mark.parametrize(
    "source",
    [
        "Text ^bad_id\n",
        "```text\nbody\n``` ^bad-code\n",
        "```text\rbody\r```\r",
    ],
)
def test_carrier_bridge_fails_closed_on_source_oracle_errors(source: str) -> None:
    with pytest.raises(RuntimeSemanticsV3Unsupported):
        prepare_resolved_source_carriers_v4(
            source,
            input_id="probe.md",
            expected_source_sha256=_sha256(source),
        )


def test_carrier_bridge_has_no_markers_for_authored_manual_heading_number() -> None:
    source = "## 第二章 2.3 标题\n"
    plan = prepare_resolved_source_carriers_v4(
        source,
        input_id="manual.md",
        expected_source_sha256=_sha256(source),
    )
    assert plan.marker_edits == ()
    assert plan.shielded_source == source
