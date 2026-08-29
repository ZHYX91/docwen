"""Fresh v4 range-marker and AST ownership gates."""

from __future__ import annotations

import copy
import hashlib
from dataclasses import replace
from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.resolved_numbering import (
    NumberingExportPlanEnvelope,
    NumberingTarget,
    ResolvedCitation,
    ResolvedCitationItem,
    ResolvedDocument,
    ResolvedDocumentEnvelope,
    ResolvedDocumentTarget,
    ResolvedNumberingPlan,
    ResolvedNumberingPort,
    ResolvedReference,
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
_REFERENCE_KEY = "_docwen_resolved_v4_reference"
_CITATION_KEY = "_docwen_resolved_v4_citation"
_CAPTION_SOURCES = {
    "figure": "Figure: Composite",
    "table": "Table: Composite",
    "equation": "Equation: Composite",
    "code_block": "Code: Composite",
}
_CARRIER_SOURCES = {
    "figure": "![image](image.png)",
    "table": "| A |\n|---|\n| 1 |",
    "equation": "$$x=1$$",
    "code_block": "```text\nx\n```",
}


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


def _caption_range(source: str, declaration: str, *, occurrence: int = 0) -> tuple[int, int]:
    start = -1
    for _index in range(occurrence + 1):
        start = source.index(declaration, start + 1)
    declaration_object_boundary = source.index("\n\n", start)
    end = source.find("\n\n", declaration_object_boundary + 2)
    return start, len(source) if end < 0 else end


def _citation(source: str, token: str, *, cluster_id: str, form: str) -> ResolvedCitation:
    start = source.index(token)
    keys = tuple(part.strip()[1:] for part in token[1:-1].split(";")) if form == "parenthetical" else (token[1:],)
    return ResolvedCitation(
        source_start=start,
        source_end=start + len(token),
        source_slice_sha256=_sha(token),
        authored_token=token,
        form=form,  # type: ignore[arg-type]
        cluster_id=cluster_id,
        items=tuple(
            ResolvedCitationItem(
                citation_key=key,
                record_id=f"record:{key}",
                record_sha256=_sha(f"record:{key}"),
                presentation=key.title(),
            )
            for key in keys
        ),
        cached_result=f"[{'; '.join(keys)}]",
    )


def _port(
    source: str,
    targets: tuple[ResolvedDocumentTarget, ...],
    *,
    references: tuple[ResolvedReference, ...] = (),
    citations: tuple[ResolvedCitation, ...] = (),
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
                references=references,
                resource_occurrences=(),
                citations=citations,
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


def _full_port() -> ResolvedNumberingPort:
    source = (
        "## 2.3 **手写标题** ^head-1\n\n"
        "Figure: 图 @cite ^fig-1\n\n"
        "![x](image.png) ^raw-image\n\n\n"
        "Table: 表\n\n"
        "| A |\n|---|\n| 1 |\n\n\n"
        "Equation: 方程\n\n"
        "$$x=1$$\n\n\n"
        "Code: 代码\n\n"
        "```python\nx = 1\n```\n\n"
        "See @[[#^head-1|标题]] and [@smith; @wang].\n"
    )
    heading_end = source.index("\n")
    targets = [
        _target(
            source,
            0,
            heading_end,
            kind="heading",
            target_id="head-1",
            heading_level=2,
            authored_text="2.3 手写标题",
        )
    ]
    for declaration, kind, target_id, title in (
        ("Figure: 图 @cite ^fig-1", "figure", "fig-1", "图 @cite"),
        ("Table: 表", "table", None, "表"),
        ("Equation: 方程", "equation", None, "方程"),
        ("Code: 代码", "code_block", None, "代码"),
    ):
        start, end = _caption_range(source, declaration)
        targets.append(
            _target(
                source,
                start,
                end,
                kind=kind,
                target_id=target_id,
                authored_text=title,
            )
        )
    targets.sort(key=lambda item: item.occurrence_key)

    reference_token = "@[[#^head-1|标题]]"
    reference_start = source.index(reference_token)
    reference = ResolvedReference(
        source_start=reference_start,
        source_end=reference_start + len(reference_token),
        source_slice_sha256=_sha(reference_token),
        authored_token=reference_token,
        target_source_start=0,
        target_source_end=heading_end,
        target_kind="heading",
        target_id="head-1",
        cached_number="1.1",
        alias="标题",
    )
    citations = (
        _citation(source, "@cite", cluster_id="cluster:caption", form="narrative"),
        _citation(source, "[@smith; @wang]", cluster_id="cluster:body", form="parenthetical"),
    )
    return _port(source, tuple(targets), references=(reference,), citations=citations)


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
        elif node.get("type") in {"semantic_cross_reference", "semantic_citation"}:
            output.append(str(node.get("raw", "")))
        children = node.get("children")
        if isinstance(children, list):
            output.append(_inline_text(children))
    return "".join(output)


@pytest.mark.parametrize("caption_first", [True, False])
@pytest.mark.parametrize("blank_lines", [0, 1])
@pytest.mark.parametrize("caption_kind", tuple(_CAPTION_SOURCES))
@pytest.mark.parametrize("carrier_kind", tuple(_CARRIER_SOURCES))
def test_cross_type_captions_bind_all_carriers_in_both_orders_and_supported_spacing(
    caption_first: bool,
    blank_lines: int,
    caption_kind: str,
    carrier_kind: str,
) -> None:
    declaration = _CAPTION_SOURCES[caption_kind]
    carrier = _CARRIER_SOURCES[carrier_kind]
    separator = "\n" * (blank_lines + 1)
    source = f"{declaration}{separator}{carrier}\n" if caption_first else f"{carrier}{separator}{declaration}\n"
    declaration_start = source.index(declaration)
    target = _target(
        source,
        declaration_start,
        declaration_start + len(declaration),
        kind=caption_kind,
        target_id=None,
        authored_text="Composite",
    )

    plan = prepare_resolved_runtime_v4(_port(source, (target,)))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    owners = [node for node in _walk(restored) if node.get(_TARGET_KEY) == target]
    assert len(owners) == 1
    assert (
        owners[0]["type"]
        == {
            "figure": "paragraph",
            "table": "table",
            "equation": "block_math",
            "code_block": "block_code",
        }[carrier_kind]
    )
    assert _inline_text(owners[0][_CAPTION_CHILDREN_KEY]) == "Composite"


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

    with pytest.raises(ResolvedRuntimeV4Unsupported, match="one unique object matching"):
        apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)


def test_exact_markers_bind_all_targets_references_and_citations_without_deriving_numbers() -> None:
    port = _full_port()
    plan = prepare_resolved_runtime_v4(port)
    raw_ast = parse_markdown_text(plan.shielded_source)
    original_ast = copy.deepcopy(raw_ast)
    restored = apply_resolved_runtime_v4(raw_ast, plan)

    assert raw_ast == original_ast
    assert len(plan.marker_edits) == 8
    assert all(edit.original != "2.3" for edit in plan.marker_edits)
    assert "2.3" in plan.shielded_source

    bound = [node for node in _walk(restored) if isinstance(node.get(_TARGET_KEY), ResolvedDocumentTarget)]
    assert [node[_TARGET_KEY].kind for node in bound] == [
        "heading",
        "figure",
        "table",
        "equation",
        "code_block",
    ]
    heading = next(node for node in bound if node[_TARGET_KEY].kind == "heading")
    assert _inline_text(heading["children"]) == "2.3 手写标题"
    assert "^head-1" not in _inline_text(heading["children"])

    figure = next(node for node in bound if node[_TARGET_KEY].kind == "figure")
    caption_nodes = list(_walk(figure[_CAPTION_CHILDREN_KEY]))
    citation_node = next(node for node in caption_nodes if node.get("type") == "semantic_citation")
    assert citation_node["raw"] == "@cite"
    assert citation_node[_CITATION_KEY].cluster_id == "cluster:caption"
    assert _inline_text(figure[_CAPTION_CHILDREN_KEY]) == "图 @cite"

    reference_node = next(node for node in _walk(restored) if _REFERENCE_KEY in node)
    assert reference_node["schema"] == "docwen.resolved_document.v1"
    assert reference_node[_REFERENCE_KEY].target_id == "head-1"
    body_citation = next(
        node for node in _walk(restored) if _CITATION_KEY in node and node[_CITATION_KEY].cluster_id == "cluster:body"
    )
    assert body_citation["raw"] == "[@smith; @wang]"


def test_idless_heading_closing_marks_and_nested_case_insensitive_caption_bind_structurally() -> None:
    source = "# **2.3 标题** @key #\n\n> figure: Nested\n>\n> ![x](image.png)\n"
    heading_end = source.index("\n")
    figure_start = source.index("> figure:")
    targets = (
        _target(
            source,
            0,
            heading_end,
            kind="heading",
            target_id=None,
            heading_level=1,
            authored_text="2.3 标题 @key",
        ),
        _target(
            source,
            figure_start,
            len(source),
            kind="figure",
            target_id=None,
            authored_text="Nested",
        ),
    )
    citation = _citation(source, "@key", cluster_id="cluster:heading", form="narrative")
    plan = prepare_resolved_runtime_v4(_port(source, targets, citations=(citation,)))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    bound = [node for node in _walk(restored) if isinstance(node.get(_TARGET_KEY), ResolvedDocumentTarget)]
    assert [node[_TARGET_KEY].kind for node in bound] == ["heading", "figure"]
    heading = bound[0]
    assert _inline_text(heading["children"]) == "2.3 标题 @key"
    assert any(_CITATION_KEY in node for node in _walk(heading["children"]))
    assert _inline_text(bound[1][_CAPTION_CHILDREN_KEY]) == "Nested"


def test_caption_first_inline_markers_compose_at_the_same_authenticated_boundary() -> None:
    source = "# Heading ^h\n\nFigure:@cite\n\n![x](image.png)\n\n\nTable:@[[#^h]]\n\n| A |\n|---|\n| 1 |\n"
    heading_end = source.index("\n")
    figure_start, figure_end = _caption_range(source, "Figure:@cite")
    table_start, table_end = _caption_range(source, "Table:@[[#^h]]")
    targets = (
        _target(
            source,
            0,
            heading_end,
            kind="heading",
            target_id="h",
            heading_level=1,
            authored_text="Heading",
        ),
        _target(
            source,
            figure_start,
            figure_end,
            kind="figure",
            target_id=None,
            authored_text="@cite",
        ),
        _target(
            source,
            table_start,
            table_end,
            kind="table",
            target_id=None,
            authored_text="@[[#^h]]",
        ),
    )
    reference_token = "@[[#^h]]"
    reference_start = source.index(reference_token)
    reference = ResolvedReference(
        source_start=reference_start,
        source_end=reference_start + len(reference_token),
        source_slice_sha256=_sha(reference_token),
        authored_token=reference_token,
        target_source_start=0,
        target_source_end=heading_end,
        target_kind="heading",
        target_id="h",
        cached_number="1",
        alias=None,
    )
    citation = _citation(source, "@cite", cluster_id="cluster:first", form="narrative")
    plan = prepare_resolved_runtime_v4(_port(source, targets, references=(reference,), citations=(citation,)))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    caption_objects = [node for node in _walk(restored) if _CAPTION_CHILDREN_KEY in node]
    assert [_inline_text(node[_CAPTION_CHILDREN_KEY]) for node in caption_objects] == [
        "@cite",
        "@[[#^h]]",
    ]

    figure_marker = next(
        item
        for item in plan.markers
        if item.role == "target" and isinstance(item.payload, ResolvedDocumentTarget) and item.payload.kind == "figure"
    )
    citation_marker = next(item for item in plan.markers if item.role == "citation")
    temporary = "DOCWENV4SWAPTEMP"
    swapped = plan.shielded_source.replace(figure_marker.marker, temporary, 1)
    swapped = swapped.replace(citation_marker.marker, figure_marker.marker, 1)
    swapped = swapped.replace(temporary, citation_marker.marker, 1)
    with pytest.raises(ResolvedRuntimeV4Unsupported, match=r"kind colon|source order"):
        apply_resolved_runtime_v4(parse_markdown_text(swapped), plan)


def test_same_kind_target_markers_cannot_swap_authenticated_source_slots() -> None:
    source = "Figure: Same\n\n![a](a.png)\n\n\nFigure: Same\n\n![b](b.png)\n"
    targets = tuple(
        _target(
            source,
            *_caption_range(source, declaration, occurrence=index),
            kind="figure",
            target_id=None,
            authored_text="Same",
        )
        for index, declaration in enumerate(("Figure: Same", "Figure: Same"))
    )
    plan = prepare_resolved_runtime_v4(_port(source, targets))
    target_markers = [item.marker for item in plan.markers if item.role == "target"]
    assert len(target_markers) == 2
    temporary = "DOCWENV4TARGETSWAPTEMP"
    swapped = plan.shielded_source.replace(target_markers[0], temporary, 1)
    swapped = swapped.replace(target_markers[1], target_markers[0], 1)
    swapped = swapped.replace(temporary, target_markers[1], 1)

    with pytest.raises(ResolvedRuntimeV4Unsupported, match="target order"):
        apply_resolved_runtime_v4(parse_markdown_text(swapped), plan)


@pytest.mark.parametrize(
    ("source", "kind", "target_id", "heading_level", "message"),
    [
        (
            "Figure: Caption\n\n![x](image.png) ^figure\n",
            "figure",
            "figure",
            None,
            "target ID",
        ),
        ("Figure: Caption\n\n![x](image.png)\n", "table", None, None, "declaration kind"),
        ("Listing: Caption\n\n```text\nx\n```\n", "code_block", None, None, "declaration kind"),
        ("List: Caption\n\n```text\nx\n```\n", "code_block", None, None, "declaration kind"),
        (": Caption\n\n![x](image.png)\n", "figure", None, None, "declaration kind"),
        ("Figure: Caption {#figure}\n\n![x](image.png)\n", "figure", None, None, "historical"),
        ("## Heading\n", "heading", None, 1, "ATX level"),
    ],
)
def test_target_marker_preparation_rejects_identity_or_kind_guessing(
    source: str,
    kind: str,
    target_id: str | None,
    heading_level: int | None,
    message: str,
) -> None:
    target = _target(
        source,
        0,
        len(source.rstrip("\n")),
        kind=kind,
        target_id=target_id,
        heading_level=heading_level,
        authored_text="Caption" if kind != "heading" else "Heading",
    )
    with pytest.raises(ResolvedRuntimeV4Unsupported, match=message):
        prepare_resolved_runtime_v4(_port(source, (target,)))


@pytest.mark.parametrize(
    ("source", "kind", "target_id", "message"),
    [
        ("Figure:\n\n![x](image.png)\n", "figure", None, "require authored content"),
        ("Table:\n\n| A |\n|---|\n| 1 |\n", "table", None, "require authored content"),
        ("Code:\n\n```text\nx\n```\n", "code_block", None, "requires an explicit source ID"),
        ("Equation:\n\n$$x=1$$\n", "equation", None, "requires an explicit source ID"),
    ],
)
def test_empty_caption_rules_fail_closed_before_ast_binding(
    source: str,
    kind: str,
    target_id: str | None,
    message: str,
) -> None:
    target = _target(
        source,
        0,
        len(source),
        kind=kind,
        target_id=target_id,
        authored_text="",
    )
    with pytest.raises(ResolvedRuntimeV4Unsupported, match=message):
        prepare_resolved_runtime_v4(_port(source, (target,)))


def test_empty_equation_with_explicit_id_remains_valid() -> None:
    source = "Equation: ^energy\n\n$$x=1$$\n"
    target = _target(
        source,
        0,
        len(source),
        kind="equation",
        target_id="energy",
        authored_text="",
    )
    plan = prepare_resolved_runtime_v4(_port(source, (target,)))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)
    equation = next(node for node in _walk(restored) if node.get(_TARGET_KEY) == target)
    assert equation[_CAPTION_CHILDREN_KEY] == []


@pytest.mark.parametrize("target_kind", ["heading", "figure"])
def test_apply_rejects_forged_authored_text(target_kind: str) -> None:
    port = _full_port()
    targets = tuple(
        replace(target, authored_text="伪造") if target.kind == target_kind else target
        for target in port.document.targets
    )
    changed_document = replace(port.document, targets=targets)
    changed_envelope = replace(port.document_envelope, document=changed_document)
    changed_port = ResolvedNumberingPort(changed_envelope, port.plan_envelope)
    plan = prepare_resolved_runtime_v4(changed_port)
    with pytest.raises(ResolvedRuntimeV4Unsupported, match="authored_text"):
        apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)


def test_markdown_heading_level_seven_binds_to_the_extended_atx_ast() -> None:
    source = "####### Heading\n"
    target = _target(
        source,
        0,
        len(source.rstrip("\n")),
        kind="heading",
        target_id=None,
        heading_level=7,
        authored_text="Heading",
    )
    plan = prepare_resolved_runtime_v4(_port(source, (target,)))
    restored = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)
    heading = next(node for node in _walk(restored) if node.get(_TARGET_KEY) == target)

    assert heading["type"] == "heading"
    assert heading["attrs"]["level"] == 7


def test_prepare_rejects_changed_source_hash_and_plan_inventory() -> None:
    port = _full_port()
    changed_document = replace(port.document_envelope, source_sha256="0" * 64)
    with pytest.raises(ResolvedRuntimeV4Unsupported, match="source hash"):
        prepare_resolved_runtime_v4(ResolvedNumberingPort(changed_document, port.plan_envelope))

    changed_plan = replace(port.plan, targets=port.plan.targets[:-1])
    changed_envelope = replace(port.plan_envelope, plan=changed_plan)
    with pytest.raises(ResolvedRuntimeV4Unsupported, match="target inventories"):
        prepare_resolved_runtime_v4(ResolvedNumberingPort(port.document_envelope, changed_envelope))


@pytest.mark.parametrize("mutation", ["missing", "duplicate"])
def test_apply_rejects_missing_or_duplicated_marker(mutation: str) -> None:
    plan = prepare_resolved_runtime_v4(_full_port())
    marker = plan.markers[0].marker
    replacement = "BROKEN" if mutation == "missing" else marker + marker
    tampered = plan.shielded_source.replace(marker, replacement, 1)
    with pytest.raises(ResolvedRuntimeV4Unsupported):
        apply_resolved_runtime_v4(parse_markdown_text(tampered), plan)


def test_caption_marker_rejects_wrong_adjacent_object_before_any_renderer_binding() -> None:
    source = "Figure: Caption\n\nnot an image\n"
    target = _target(
        source,
        0,
        len(source),
        kind="figure",
        target_id=None,
        authored_text="Caption",
    )
    plan = prepare_resolved_runtime_v4(_port(source, (target,)))
    with pytest.raises(ResolvedRuntimeV4Unsupported, match="one unique object matching"):
        apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)


def test_next_line_heading_and_caption_ids_bind_to_targets_not_carriers() -> None:
    source = "# Heading\n^heading-id\n\nFigure: Composite\n^caption-id\n| A |\n|---|\n| 1 |\n"
    heading_end = len("# Heading")
    caption_start = source.index("Figure:")
    caption_end = caption_start + len("Figure: Composite")
    targets = (
        _target(
            source,
            0,
            heading_end,
            kind="heading",
            target_id="heading-id",
            heading_level=1,
            authored_text="Heading",
        ),
        _target(
            source,
            caption_start,
            caption_end,
            kind="figure",
            target_id="caption-id",
            authored_text="Composite",
        ),
    )

    plan = prepare_resolved_runtime_v4(_port(source, targets))
    ast = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    heading, table = [node for node in ast if node.get("type") != "blank_line"]
    assert heading[_TARGET_KEY].target_id == "heading-id"
    assert table["type"] == "table"
    assert table[_TARGET_KEY].target_id == "caption-id"
    assert all("^heading-id" not in str(node) and "^caption-id" not in str(node) for node in ast)


@pytest.mark.parametrize("declaration", ["Code: ^snippet", "Code:\n^snippet"])
def test_empty_code_with_explicit_id_is_valid(declaration: str) -> None:
    source = f"{declaration}\n```text\nx\n```\n"
    target = _target(
        source,
        0,
        len(declaration.splitlines()[0]),
        kind="code_block",
        target_id="snippet",
        authored_text="",
    )

    plan = prepare_resolved_runtime_v4(_port(source, (target,)))
    [node] = apply_resolved_runtime_v4(parse_markdown_text(plan.shielded_source), plan)

    assert node["type"] == "block_code"
    assert node[_TARGET_KEY].target_id == "snippet"


def test_module_has_no_v3_or_legacy_numbering_parser_dependency() -> None:
    source_root = Path(__file__).parents[1].joinpath("src/docwen_plugin_markdown")
    module_source = source_root.joinpath("resolved_runtime_v4.py").read_text(encoding="utf-8")
    module_source += source_root.joinpath("_resolved_runtime_v4_evidence.py").read_text(encoding="utf-8")
    forbidden = (
        "document_semantics_v3",
        "runtime_semantics_v3",
        "heading_counters",
        "add_md_numbering",
        "remove_md_numbering",
    )
    assert all(name not in module_source for name in forbidden)
