"""Resolved-v4 standalone caption and canonical keyword gates."""

from __future__ import annotations

import hashlib
from typing import Any

import pytest

from docwen_core.models.resolved_numbering import (
    NumberingExportPlanEnvelope,
    NumberingTarget,
    ResolvedDocument,
    ResolvedDocumentEnvelope,
    ResolvedDocumentTarget,
    ResolvedNumberingPlan,
    ResolvedNumberingPort,
)
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.resolved_runtime_v4 import (
    ResolvedRuntimeV4Unsupported,
    apply_resolved_runtime_v4,
    prepare_resolved_runtime_v4,
)

pytestmark = pytest.mark.unit

_TARGET_KEY = "_docwen_resolved_v4_target"
_CAPTION_CHILDREN_KEY = "_docwen_resolved_v4_caption_children"


def _sha(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _target(
    source: str,
    start: int,
    end: int,
    *,
    kind: str,
    target_id: str | None,
    authored_text: str,
    heading_level: int | None = None,
) -> ResolvedDocumentTarget:
    return ResolvedDocumentTarget(
        source_start=start,
        source_end=end,
        source_slice_sha256=_sha(source[start:end]),
        kind=kind,  # type: ignore[arg-type]
        target_id=target_id,
        heading_level=heading_level,
        authored_text=authored_text,
    )


def _port(
    source: str,
    targets: tuple[ResolvedDocumentTarget, ...],
) -> ResolvedNumberingPort:
    plan_sha256 = "a" * 64
    plan_targets = tuple(
        NumberingTarget(
            source_start=target.source_start,
            source_end=target.source_end,
            kind=target.kind,
            enabled=False,
            target_id=target.target_id,
            derived_number=None,
            materialization=None,
        )
        for target in targets
    )
    return ResolvedNumberingPort(
        ResolvedDocumentEnvelope(
            input_id="neutral",
            source_sha256=_sha(source),
            plan_sha256=plan_sha256,
            document=ResolvedDocument(
                authored_markdown=source,
                targets=targets,
                references=(),
                resource_occurrences=(),
                citations=(),
                resources=(),
            ),
        ),
        NumberingExportPlanEnvelope(
            input_id="neutral",
            source_sha256=_sha(source),
            plan_sha256=plan_sha256,
            plan=ResolvedNumberingPlan(
                heading_definitions=(),
                heading_instances=(),
                targets=plan_targets,
            ),
        ),
    )


def _walk(nodes: list[dict[str, Any]]):
    for node in nodes:
        yield node
        children = node.get("children")
        if isinstance(children, list):
            yield from _walk(children)
        caption_children = node.get(_CAPTION_CHILDREN_KEY)
        if isinstance(caption_children, list):
            yield from _walk(caption_children)


def _inline_text(nodes: list[dict[str, Any]]) -> str:
    output: list[str] = []
    for node in nodes:
        if node.get("type") == "text":
            output.append(str(node.get("raw", node.get("text", ""))))
        children = node.get("children")
        if isinstance(children, list):
            output.append(_inline_text(children))
    return "".join(output)


def test_caption_without_captionable_neighbor_stays_standalone() -> None:
    source = "Figure: Caption\n\nnot an image\n"
    target = _target(
        source,
        0,
        len("Figure: Caption"),
        kind="figure",
        target_id=None,
        authored_text="Caption",
    )
    plan = prepare_resolved_runtime_v4(_port(source, (target,)))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    owner = next(node for node in _walk(restored) if node.get(_TARGET_KEY) == target)
    assert owner["type"] == "_docwen_resolved_v4_caption_declaration"
    assert _inline_text(owner[_CAPTION_CHILDREN_KEY]) == "Caption"


def test_resolved_chain_does_not_use_global_matching_to_resolve_local_ambiguity() -> None:
    source = "Figure: First\n\n| first |\n|---|\n| 1 |\n\nTable: Second\n\n![second](second.png)\n"
    targets = tuple(
        _target(
            source,
            source.index(declaration),
            source.index(declaration) + len(declaration),
            kind=kind,
            target_id=None,
            authored_text=title,
        )
        for declaration, kind, title in (
            ("Figure: First", "figure", "First"),
            ("Table: Second", "table", "Second"),
        )
    )
    plan = prepare_resolved_runtime_v4(_port(source, targets))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    bound = [node for node in _walk(restored) if node.get(_TARGET_KEY) == targets[0]]
    standalone = [node for node in _walk(restored) if node.get(_TARGET_KEY) == targets[1]]
    assert len(bound) == 1 and bound[0]["type"] == "table"
    assert len(standalone) == 1
    assert standalone[0]["type"] == "_docwen_resolved_v4_caption_declaration"
    assert _inline_text(standalone[0][_CAPTION_CHILDREN_KEY]) == "Second"


def test_two_sided_caption_ambiguity_stays_standalone() -> None:
    source = "![before](before.png)\n\nFigure: Comparison\n\n| A |\n|---|\n| 1 |\n"
    start = source.index("Figure:")
    target = _target(
        source,
        start,
        start + len("Figure: Comparison"),
        kind="figure",
        target_id=None,
        authored_text="Comparison",
    )
    plan = prepare_resolved_runtime_v4(_port(source, (target,)))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    owner = next(node for node in _walk(restored) if node.get(_TARGET_KEY) == target)
    assert owner["type"] == "_docwen_resolved_v4_caption_declaration"


def test_two_captions_competing_for_one_carrier_both_stay_standalone() -> None:
    source = "Figure: Above\n\n![shared](shared.png)\n\nTable: Below\n"
    targets = tuple(
        _target(
            source,
            source.index(declaration),
            source.index(declaration) + len(declaration),
            kind=kind,
            target_id=None,
            authored_text=title,
        )
        for declaration, kind, title in (
            ("Figure: Above", "figure", "Above"),
            ("Table: Below", "table", "Below"),
        )
    )
    plan = prepare_resolved_runtime_v4(_port(source, targets))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    owners = [node for node in _walk(restored) if node.get(_TARGET_KEY) in targets]
    assert len(owners) == 2
    assert all(node["type"] == "_docwen_resolved_v4_caption_declaration" for node in owners)


def test_lowercase_caption_keyword_is_not_an_interop_declaration() -> None:
    source = "figure: Nested\n\n![x](image.png)\n"
    target = _target(
        source,
        0,
        len("figure: Nested"),
        kind="figure",
        target_id=None,
        authored_text="Nested",
    )

    with pytest.raises(
        ResolvedRuntimeV4Unsupported,
        match="does not contain one exact declaration kind",
    ):
        prepare_resolved_runtime_v4(_port(source, (target,)))
