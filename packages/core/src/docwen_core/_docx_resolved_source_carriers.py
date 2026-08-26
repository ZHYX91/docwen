"""Source-authenticated block-carrier ordering for resolved v4 DOCX output.

This module deliberately owns no numbering or reference semantics.  It admits
only the frozen ordinary-anchor carrier record, combines that record with the
already-validated resolved target/occurrence groups, and derives one laminar
physical wrapper order.  Equal physical ranges are directed only by closed
source ownership facts; they are never resolved by bind order or visible text.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from itertools import combinations, pairwise
from typing import Any, Literal

from docwen_core._docx_semantics_v3_model import (
    ANCHOR_TAG_PREFIX,
    TARGET_TAG_PREFIX,
    AnchorIdentityV3,
    AnchorTopologyEdgeV3,
    DocxSemanticsV3Error,
    derive_anchor_identity_v3,
    derive_anchor_topology_edge_v3,
    require_source_id,
)
from docwen_core._docx_semantics_v3_ooxml import sdt_tag

_NUMBERING_OCCURRENCE_TAG_PREFIX = "docwen-numbering-occurrence-v1:"
_BLOCK_TAG_PREFIXES = (TARGET_TAG_PREFIX, ANCHOR_TAG_PREFIX, _NUMBERING_OCCURRENCE_TAG_PREFIX)
_CONTAINER_KINDS = frozenset({"list", "list_item", "block_quote", "callout"})
_SOURCE_BLOCK_KINDS = frozenset(
    {
        "paragraph",
        "image",
        "table",
        "equation",
        "code_block",
        "fenced_block",
        "list",
        "list_item",
        "block_quote",
        "callout",
    }
)
_RAW_TARGET_KIND = {
    "image": "figure",
    "table": "table",
    "equation": "equation",
    "code_block": "code_block",
    "fenced_block": "code_block",
}
_POST_BLOCK_GAP_CHARS = frozenset(" \t\r\n>+-*0123456789.)[]xX")

type ResolvedCarrierRole = Literal["anchor", "target", "occurrence"]


@dataclass(frozen=True, slots=True)
class SourceOwnerRangeV4:
    block_kind: str
    source_start: int
    source_end: int


@dataclass(frozen=True, slots=True)
class ResolvedOrdinaryAnchorBindingV4:
    identity: AnchorIdentityV3
    elements: tuple[Any, ...]
    direct_parent_source_id: str | None
    source_block_kind: str
    source_start: int
    source_end: int
    marker_start: int
    marker_end: int
    owner_path: tuple[SourceOwnerRangeV4, ...]


@dataclass(frozen=True, slots=True)
class ResolvedBlockCarrierGroupV4:
    role: ResolvedCarrierRole
    tag: str
    elements: tuple[Any, ...]
    source_start: int
    source_end: int
    source_kind: str
    payload: Any


@dataclass(frozen=True, slots=True)
class ResolvedCarrierOrderV4:
    ordered: tuple[ResolvedBlockCarrierGroupV4, ...]
    direct_parent_by_tag: dict[str, str | None]
    anchor_topology_edges: tuple[AnchorTopologyEdgeV3, ...]


@dataclass(frozen=True, slots=True)
class _CarrierRow:
    group: ResolvedBlockCarrierGroupV4
    body_start: int
    body_end: int


def admit_resolved_ordinary_anchor_v4(
    authored_source: str,
    elements: tuple[Any, ...],
    anchor: Mapping[str, Any],
    *,
    direct_parent_source_id: str | None,
) -> ResolvedOrdinaryAnchorBindingV4:
    """Admit one closed source-oracle anchor record without parsing Markdown."""

    expected_fields = {"id", "block_kind", "placement", "range", "block_range", "container_path"}
    if set(anchor) != expected_fields:
        raise DocxSemanticsV3Error("resolved ordinary-anchor fields are not closed")
    source_id = _required_text(anchor["id"], "ordinary-anchor id")
    require_source_id(source_id)
    source_kind = _required_source_kind(anchor["block_kind"], "ordinary-anchor block kind")
    placement = _required_text(anchor["placement"], "ordinary-anchor placement")
    if placement not in {"inline", "post_block"}:
        raise DocxSemanticsV3Error("resolved ordinary-anchor placement is not closed")
    marker_start, marker_end = _closed_range(anchor["range"], authored_source, "ordinary-anchor marker")
    block_start, block_end = _closed_range(anchor["block_range"], authored_source, "ordinary-anchor block")
    if authored_source[marker_start:marker_end] != f"^{source_id}":
        raise DocxSemanticsV3Error("ordinary-anchor marker range differs from its source ID")
    if placement == "inline":
        if not (block_start <= marker_start < marker_end <= block_end):
            raise DocxSemanticsV3Error("inline ordinary-anchor marker is outside its source block")
    else:
        gap = authored_source[block_end:marker_start]
        if block_end > marker_start or any(character not in _POST_BLOCK_GAP_CHARS for character in gap):
            raise DocxSemanticsV3Error("post-block ordinary-anchor marker has a non-container source gap")

    container_path = anchor["container_path"]
    if not isinstance(container_path, list):
        raise DocxSemanticsV3Error("ordinary-anchor container_path must be an array")
    owners: list[SourceOwnerRangeV4] = []
    for index, item in enumerate(container_path):
        if not isinstance(item, Mapping) or set(item) != {"block_kind", "block_range"}:
            raise DocxSemanticsV3Error("ordinary-anchor container_path segment is not closed")
        kind = _required_source_kind(item["block_kind"], "ordinary-anchor container kind")
        if kind not in _CONTAINER_KINDS:
            raise DocxSemanticsV3Error("ordinary-anchor container_path contains a non-container kind")
        start, end = _closed_range(
            item["block_range"],
            authored_source,
            f"ordinary-anchor container_path[{index}]",
        )
        owners.append(SourceOwnerRangeV4(kind, start, end))
    owners.append(SourceOwnerRangeV4(source_kind, block_start, block_end))
    for outer, inner in pairwise(owners):
        if not (outer.source_start <= inner.source_start and inner.source_end <= outer.source_end and outer != inner):
            raise DocxSemanticsV3Error("ordinary-anchor source owner path is not a proper nested path")
    if not elements or len({id(item) for item in elements}) != len(elements):
        raise DocxSemanticsV3Error("ordinary-anchor group must contain distinct elements")
    if direct_parent_source_id is not None:
        require_source_id(direct_parent_source_id)
        if direct_parent_source_id == source_id:
            raise DocxSemanticsV3Error("ordinary-anchor direct parent must differ from its child")

    identity_kind = "code_block" if source_kind == "fenced_block" else source_kind
    return ResolvedOrdinaryAnchorBindingV4(
        identity=derive_anchor_identity_v3(identity_kind, source_id),
        elements=elements,
        direct_parent_source_id=direct_parent_source_id,
        source_block_kind=source_kind,
        source_start=block_start,
        source_end=block_end,
        marker_start=marker_start,
        marker_end=marker_end,
        owner_path=tuple(owners),
    )


def resolved_anchor_group_v4(binding: ResolvedOrdinaryAnchorBindingV4) -> ResolvedBlockCarrierGroupV4:
    return ResolvedBlockCarrierGroupV4(
        role="anchor",
        tag=binding.identity.tag,
        elements=binding.elements,
        source_start=binding.source_start,
        source_end=binding.source_end,
        source_kind=binding.source_block_kind,
        payload=binding,
    )


def resolved_semantic_group_v4(
    *,
    role: Literal["target", "occurrence"],
    tag: str,
    elements: tuple[Any, ...],
    source_start: int,
    source_end: int,
    source_kind: str,
    payload: Any,
) -> ResolvedBlockCarrierGroupV4:
    if source_kind not in {"heading", "figure", "table", "equation", "code_block"}:
        raise DocxSemanticsV3Error("resolved semantic carrier kind is outside the closed set")
    expected_prefix = TARGET_TAG_PREFIX if role == "target" else _NUMBERING_OCCURRENCE_TAG_PREFIX
    if not tag.startswith(expected_prefix):
        raise DocxSemanticsV3Error("resolved semantic carrier tag has the wrong role prefix")
    if source_start < 0 or source_end <= source_start or not elements:
        raise DocxSemanticsV3Error("resolved semantic carrier range/group is invalid")
    return ResolvedBlockCarrierGroupV4(
        role=role,
        tag=tag,
        elements=elements,
        source_start=source_start,
        source_end=source_end,
        source_kind=source_kind,
        payload=payload,
    )


def order_resolved_block_carriers_v4(
    body: Any,
    groups: tuple[ResolvedBlockCarrierGroupV4, ...],
    anchors: tuple[ResolvedOrdinaryAnchorBindingV4, ...],
) -> ResolvedCarrierOrderV4:
    """Prove one laminar family and return its deterministic child-first order."""

    anchor_parent_by_tag, edges = _prove_anchor_source_forest(anchors)
    if len({item.tag for item in groups}) != len(groups):
        raise DocxSemanticsV3Error("resolved block-carrier tags are not unique")
    rows: dict[str, _CarrierRow] = {}
    for group in groups:
        if not group.elements or len({id(item) for item in group.elements}) != len(group.elements):
            raise DocxSemanticsV3Error("resolved block-carrier group must contain distinct elements")
        if any(item.getparent() is not body for item in group.elements):
            raise DocxSemanticsV3Error("resolved block-carrier group requires direct main-body elements")
        positions = [body.index(item) for item in group.elements]
        if positions != list(range(positions[0], positions[0] + len(positions))):
            raise DocxSemanticsV3Error("resolved block-carrier group must be contiguous and ordered")
        rows[group.tag] = _CarrierRow(group, positions[0], positions[-1])

    relation: set[tuple[str, str]] = set()  # (inner, outer)
    for left, right in combinations(rows.values(), 2):
        direction = _physical_pair_direction(left, right, anchor_parent_by_tag)
        if direction is not None:
            relation.add(direction)

    ordered_tags: list[str] = []
    pending = set(rows)
    while pending:
        ready = [tag for tag in pending if not any(inner in pending and outer == tag for inner, outer in relation)]
        if not ready:
            raise DocxSemanticsV3Error("resolved block-carrier ownership graph contains a cycle")
        ready.sort(
            key=lambda tag: (
                rows[tag].body_start,
                rows[tag].body_end - rows[tag].body_start,
                tag,
            )
        )
        ordered_tags.extend(ready)
        pending.difference_update(ready)

    direct_parent_by_tag: dict[str, str | None] = {}
    for child in rows:
        candidates = {outer for inner, outer in relation if inner == child}
        nearest = {
            candidate
            for candidate in candidates
            if not any(other != candidate and (other, candidate) in relation for other in candidates)
        }
        if len(nearest) > 1:
            raise DocxSemanticsV3Error("resolved block-carrier has ambiguous nearest ownership")
        direct_parent_by_tag[child] = next(iter(nearest), None)
    return ResolvedCarrierOrderV4(
        ordered=tuple(rows[tag].group for tag in ordered_tags),
        direct_parent_by_tag=direct_parent_by_tag,
        anchor_topology_edges=edges,
    )


def validate_fenced_anchor_ranges_v4(
    body: Any,
    anchors: tuple[ResolvedOrdinaryAnchorBindingV4, ...],
    fenced_bindings: tuple[Any, ...],
) -> None:
    """Require source and rendered containment to agree for every fence/anchor pair."""

    anchor_rows: list[tuple[ResolvedOrdinaryAnchorBindingV4, int, int]] = []
    for anchor in anchors:
        positions = [body.index(item) for item in anchor.elements if item.getparent() is body]
        if len(positions) != len(anchor.elements):
            raise DocxSemanticsV3Error("fenced-source containment requires unwrapped anchor elements")
        anchor_rows.append((anchor, positions[0], positions[-1]))
    for fenced in fenced_bindings:
        paragraph = fenced.paragraph_element
        if paragraph.getparent() is not body:
            raise DocxSemanticsV3Error("fenced-source paragraph is not a direct main-body element")
        position = body.index(paragraph)
        identity = fenced.identity
        for anchor, start, end in anchor_rows:
            physically_inside = start <= position <= end
            source_inside = anchor.source_start <= identity.source_start and identity.source_end <= anchor.source_end
            if physically_inside != source_inside:
                raise DocxSemanticsV3Error("fenced-source source range contradicts ordinary-anchor ownership")


def physical_resolved_block_tags_v4(body: Any) -> tuple[str, ...]:
    from docx.oxml.ns import qn

    return tuple(
        tag
        for item in body.iter(qn("w:sdt"))
        if (tag := sdt_tag(item)) is not None and tag.startswith(_BLOCK_TAG_PREFIXES)
    )


def prove_resolved_block_hierarchy_v4(
    body: Any,
    *,
    expected_tags: set[str],
    expected_physical_tags: tuple[str, ...],
    expected_parent_by_tag: Mapping[str, str | None],
) -> dict[str, Any]:
    """Bind every owned block tag exactly once and prove its immediate owner."""

    from docx.oxml.ns import qn

    from docwen_core._docx_semantics_v3_topology import prove_block_sdt_envelope

    wrappers: dict[str, Any] = {}
    physical: list[str] = []
    for wrapper in body.iter(qn("w:sdt")):
        tag = sdt_tag(wrapper) or ""
        if not tag.startswith(_BLOCK_TAG_PREFIXES):
            continue
        if tag not in expected_tags or tag in wrappers:
            raise DocxSemanticsV3Error("resolved block carrier is duplicated or unmapped")
        prove_block_sdt_envelope(wrapper, tag)
        wrappers[tag] = wrapper
        physical.append(tag)
    if set(wrappers) != expected_tags or tuple(physical) != expected_physical_tags:
        raise DocxSemanticsV3Error("resolved block-carrier physical order/cardinality differs from authority")
    for tag, wrapper in wrappers.items():
        parent = wrapper.getparent()
        if parent is body:
            parent_tag = None
        elif parent is not None and parent.tag == qn("w:sdtContent"):
            parent_tag = sdt_tag(parent.getparent())
            if parent_tag not in expected_tags:
                raise DocxSemanticsV3Error("resolved block carrier has an unowned block parent")
        else:
            raise DocxSemanticsV3Error("resolved block carrier is outside the supported main-body hierarchy")
        if expected_parent_by_tag.get(tag) != parent_tag:
            raise DocxSemanticsV3Error("resolved block-carrier direct parent differs after reopen")
    return wrappers


def prove_resolved_anchor_topology_v4(
    anchors: tuple[ResolvedOrdinaryAnchorBindingV4, ...],
    records: tuple[AnchorTopologyEdgeV3, ...],
    wrappers_by_tag: Mapping[str, Any],
) -> None:
    """Bind source topology to the nearest physical ordinary ancestor."""

    anchor_tags = {item.identity.tag for item in anchors}
    if set(wrappers_by_tag).intersection(anchor_tags) != anchor_tags:
        raise DocxSemanticsV3Error("ordinary-anchor wrapper inventory is incomplete")
    physical: list[AnchorTopologyEdgeV3] = []
    for tag in anchor_tags:
        wrapper = wrappers_by_tag[tag]
        parent_tag = next(
            (
                value
                for ancestor in wrapper.iterancestors()
                if ancestor.tag.endswith("}sdt")
                and (value := sdt_tag(ancestor)) is not None
                and value.startswith(ANCHOR_TAG_PREFIX)
            ),
            None,
        )
        if parent_tag is not None:
            physical.append(derive_anchor_topology_edge_v3(tag, parent_tag))
    expected = tuple(sorted(physical, key=lambda item: (item.child_tag, item.parent_tag)))
    if records != expected:
        raise DocxSemanticsV3Error("ordinary-anchor topology differs from nearest physical nesting")


def flatten_resolved_block_elements_v4(blocks: tuple[Any, ...]) -> tuple[Any, ...]:
    """Flatten only authenticated block-carrier envelopes, never inline SDTs."""

    from docx.oxml.ns import qn

    output: list[Any] = []
    for block in blocks:
        tag = sdt_tag(block) if block.tag == qn("w:sdt") else None
        if tag is not None and tag.startswith(_BLOCK_TAG_PREFIXES):
            content = block.find(qn("w:sdtContent"))
            if content is None or not list(content):
                raise DocxSemanticsV3Error("resolved block carrier has no logical content")
            output.extend(flatten_resolved_block_elements_v4(tuple(content)))
        else:
            output.append(block)
    return tuple(output)


def prove_resolved_ordinary_anchor_group_v4(
    wrapper: Any,
    binding: ResolvedOrdinaryAnchorBindingV4,
    *,
    allowed_caption_paragraphs: set[Any],
) -> tuple[Any, ...]:
    """Prove logical kind while rejecting machinery not owned by resolved v4."""

    from docx.oxml.ns import qn

    content = wrapper.find(qn("w:sdtContent"))
    if content is None:
        raise DocxSemanticsV3Error("ordinary-anchor wrapper has no content")
    logical = flatten_resolved_block_elements_v4(tuple(content))
    kind = binding.identity.block_kind
    if kind == "table":
        valid = len(logical) == 1 and logical[0].tag == qn("w:tbl")
    elif kind == "paragraph":
        valid = len(logical) == 1 and logical[0].tag == qn("w:p")
    elif kind == "image":
        valid = (
            len(logical) == 1
            and logical[0].tag == qn("w:p")
            and (
                logical[0].find(f".//{qn('w:drawing')}") is not None
                or logical[0].find(f".//{qn('w:pict')}") is not None
            )
        )
    elif kind == "equation":
        valid = len(logical) == 1 and (
            logical[0].find(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath") is not None
            or logical[0].find(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara") is not None
        )
    elif kind in {"code_block", "list_item"}:
        valid = bool(logical) and all(item.tag == qn("w:p") for item in logical)
    elif kind in {"list", "block_quote", "callout"}:
        valid = bool(logical) and all(item.tag in {qn("w:p"), qn("w:tbl")} for item in logical)
    else:  # pragma: no cover - identity derivation owns the closed set
        valid = False
    if not valid:
        raise DocxSemanticsV3Error("ordinary-anchor group has wrong resolved block kind or cardinality")

    machine_names = ("w:bookmarkStart", "w:bookmarkEnd", "w:instrText", "w:fldChar")
    for name in machine_names:
        for item in wrapper.iter(qn(name)):
            if not _machine_item_is_resolved_owned(item, wrapper, allowed_caption_paragraphs):
                raise DocxSemanticsV3Error("ordinary anchor contains machinery not owned by resolved v4")
    return logical


def _machine_item_is_resolved_owned(item: Any, anchor_wrapper: Any, allowed_captions: set[Any]) -> bool:
    allowed_inline_prefixes = (
        TARGET_TAG_PREFIX,
        _NUMBERING_OCCURRENCE_TAG_PREFIX,
        "docwen-soft-ref-v1:",
        "docwen-ref-occurrence-v1:",
        "docwen-citation-occurrence-v1:",
    )
    current = item.getparent()
    while current is not None and current is not anchor_wrapper:
        if current in allowed_captions:
            return True
        tag = sdt_tag(current)
        if tag is not None and tag.startswith(allowed_inline_prefixes):
            return True
        current = current.getparent()
    return False


def _prove_anchor_source_forest(
    anchors: tuple[ResolvedOrdinaryAnchorBindingV4, ...],
) -> tuple[dict[str, str], tuple[AnchorTopologyEdgeV3, ...]]:
    by_source_id = {item.identity.source_id: item for item in anchors}
    if len(by_source_id) != len(anchors):
        raise DocxSemanticsV3Error("ordinary-anchor source IDs are not unique")
    if len({item.owner_path for item in anchors}) != len(anchors):
        raise DocxSemanticsV3Error("ordinary anchors duplicate one source owner path")
    parent_by_tag: dict[str, str] = {}
    edges: list[AnchorTopologyEdgeV3] = []
    for child in anchors:
        candidates = [
            parent
            for parent in anchors
            if parent is not child
            and len(parent.owner_path) < len(child.owner_path)
            and child.owner_path[: len(parent.owner_path)] == parent.owner_path
        ]
        if candidates:
            longest = max(len(item.owner_path) for item in candidates)
            nearest = [item for item in candidates if len(item.owner_path) == longest]
            if len(nearest) != 1:
                raise DocxSemanticsV3Error("ordinary-anchor source hierarchy has ambiguous direct parents")
            expected_parent = nearest[0]
            expected_source_id: str | None = expected_parent.identity.source_id
        else:
            expected_parent = None
            expected_source_id = None
        if child.direct_parent_source_id != expected_source_id:
            raise DocxSemanticsV3Error("ordinary-anchor direct parent differs from authenticated source hierarchy")
        if expected_parent is not None:
            edge = derive_anchor_topology_edge_v3(child.identity.tag, expected_parent.identity.tag)
            parent_by_tag[edge.child_tag] = edge.parent_tag
            edges.append(edge)
    return parent_by_tag, tuple(sorted(edges, key=lambda item: (item.child_tag, item.parent_tag)))


def _physical_pair_direction(
    left: _CarrierRow,
    right: _CarrierRow,
    anchor_parent_by_tag: Mapping[str, str],
) -> tuple[str, str] | None:
    if left.body_end < right.body_start or right.body_end < left.body_start:
        if _anchor_ancestor(left.group.tag, right.group.tag, anchor_parent_by_tag) or _anchor_ancestor(
            right.group.tag, left.group.tag, anchor_parent_by_tag
        ):
            raise DocxSemanticsV3Error("ordinary-anchor source parent does not physically contain its child")
        return None
    left_contains = left.body_start <= right.body_start and right.body_end <= left.body_end
    right_contains = right.body_start <= left.body_start and left.body_end <= right.body_end
    if not left_contains and not right_contains:
        raise DocxSemanticsV3Error("resolved block-carrier groups partially overlap")
    equal = left.body_start == right.body_start and left.body_end == right.body_end
    if equal:
        outer, inner = _equal_range_direction(left.group, right.group, anchor_parent_by_tag)
    elif left_contains:
        outer, inner = left.group, right.group
    else:
        outer, inner = right.group, left.group
    if not _source_allows_ownership(outer, inner, anchor_parent_by_tag):
        raise DocxSemanticsV3Error("resolved block-carrier physical nesting contradicts source authority")
    return inner.tag, outer.tag


def _equal_range_direction(
    left: ResolvedBlockCarrierGroupV4,
    right: ResolvedBlockCarrierGroupV4,
    anchor_parent_by_tag: Mapping[str, str],
) -> tuple[ResolvedBlockCarrierGroupV4, ResolvedBlockCarrierGroupV4]:
    left_outer = _source_allows_ownership(left, right, anchor_parent_by_tag)
    right_outer = _source_allows_ownership(right, left, anchor_parent_by_tag)
    if left_outer == right_outer:
        raise DocxSemanticsV3Error("equal resolved block-carrier range lacks one authenticated direction")
    return (left, right) if left_outer else (right, left)


def _source_allows_ownership(
    outer: ResolvedBlockCarrierGroupV4,
    inner: ResolvedBlockCarrierGroupV4,
    anchor_parent_by_tag: Mapping[str, str],
) -> bool:
    if outer.role == "anchor" and inner.role == "anchor":
        return _anchor_ancestor(outer.tag, inner.tag, anchor_parent_by_tag)
    if outer.role == "anchor":
        return (
            outer.source_kind in _CONTAINER_KINDS
            and outer.source_start <= inner.source_start
            and inner.source_end <= outer.source_end
        )
    if inner.role == "anchor":
        return (
            _RAW_TARGET_KIND.get(inner.source_kind) == outer.source_kind
            and outer.source_start <= inner.source_start
            and inner.source_end <= outer.source_end
        )
    return False


def _anchor_ancestor(ancestor: str, child: str, parent_by_tag: Mapping[str, str]) -> bool:
    while child in parent_by_tag:
        child = parent_by_tag[child]
        if child == ancestor:
            return True
    return False


def _closed_range(value: Any, source: str, context: str) -> tuple[int, int]:
    if not isinstance(value, Mapping) or set(value) != {"start", "end"}:
        raise DocxSemanticsV3Error(f"{context} range is not closed")
    start = value["start"]
    end = value["end"]
    if type(start) is not int or type(end) is not int or start < 0 or end <= start or end > len(source):
        raise DocxSemanticsV3Error(f"{context} range is outside authored source")
    return start, end


def _required_text(value: Any, context: str) -> str:
    if type(value) is not str or not value:
        raise DocxSemanticsV3Error(f"{context} must be non-empty text")
    return value


def _required_source_kind(value: Any, context: str) -> str:
    kind = _required_text(value, context)
    if kind not in _SOURCE_BLOCK_KINDS:
        raise DocxSemanticsV3Error(f"{context} is outside the closed set")
    return kind


__all__ = [
    "ResolvedBlockCarrierGroupV4",
    "ResolvedCarrierOrderV4",
    "ResolvedOrdinaryAnchorBindingV4",
    "admit_resolved_ordinary_anchor_v4",
    "flatten_resolved_block_elements_v4",
    "order_resolved_block_carriers_v4",
    "physical_resolved_block_tags_v4",
    "prove_resolved_anchor_topology_v4",
    "prove_resolved_block_hierarchy_v4",
    "prove_resolved_ordinary_anchor_group_v4",
    "resolved_anchor_group_v4",
    "resolved_semantic_group_v4",
    "validate_fenced_anchor_ranges_v4",
]
