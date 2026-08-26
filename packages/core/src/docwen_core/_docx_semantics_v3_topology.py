"""Closed OOXML envelope proofs for semantic v3 SDTs and REF runs."""

from __future__ import annotations

from itertools import combinations, pairwise
from typing import Any

from docwen_core._docx_semantics_v3_model import (
    ANCHOR_TAG_PREFIX,
    ANCHOR_TOPOLOGY_MAP_NAMESPACE,
    TARGET_TAG_PREFIX,
    AnchorTopologyEdgeV3,
    DocxSemanticsV3Error,
    derive_anchor_topology_edge_v3,
)


def order_ordinary_anchor_bindings(
    body: Any,
    bindings: list[Any],
) -> tuple[list[Any], tuple[AnchorTopologyEdgeV3, ...]]:
    """Prove caller-authenticated topology and order bindings child-first."""

    rows: list[tuple[Any, int, int]] = []
    for binding in bindings:
        elements = binding.elements
        if not elements or len({id(item) for item in elements}) != len(elements):
            raise DocxSemanticsV3Error("ordinary-anchor group must contain distinct elements")
        if any(item.getparent() is not body for item in elements):
            raise DocxSemanticsV3Error("ordinary-anchor group requires direct main-body elements")
        positions = [body.index(item) for item in elements]
        if positions != list(range(positions[0], positions[0] + len(positions))):
            raise DocxSemanticsV3Error("ordinary-anchor group must be contiguous and ordered")
        rows.append((binding, positions[0], positions[-1]))
    by_source_id = {row[0].identity.source_id: row for row in rows}
    parent_by_tag: dict[str, str] = {}
    edges: list[AnchorTopologyEdgeV3] = []
    for binding, _start, _end in rows:
        parent_source_id = binding.direct_parent_source_id
        if parent_source_id is None:
            continue
        parent = by_source_id.get(parent_source_id)
        if parent is None:
            raise DocxSemanticsV3Error("ordinary-anchor direct parent is not an authenticated ordinary anchor")
        edge = derive_anchor_topology_edge_v3(binding.identity.tag, parent[0].identity.tag)
        parent_by_tag[edge.child_tag] = edge.parent_tag
        edges.append(edge)
    _prove_topology_forest(edges, {row[0].identity.tag for row in rows})
    for left, right in combinations(rows, 2):
        _prove_binding_pair(left, right, parent_by_tag)

    depths = {row[0].identity.tag: _topology_depth(row[0].identity.tag, parent_by_tag) for row in rows}
    ordered = [
        row[0]
        for row in sorted(
            rows,
            key=lambda row: (
                -depths[row[0].identity.tag],
                row[2] - row[1],
                row[1],
                row[0].identity.tag,
            ),
        )
    ]
    return ordered, tuple(sorted(edges, key=lambda item: (item.child_tag, item.parent_tag)))


def anchor_topology_map_xml(records: list[AnchorTopologyEdgeV3]) -> bytes:
    """Serialize the exact independent direct-parent topology map."""

    _prove_topology_forest(records)
    if not records:
        raise DocxSemanticsV3Error("anchor-topology map requires at least one edge")
    canonical: list[AnchorTopologyEdgeV3] = []
    for item in records:
        derived = derive_anchor_topology_edge_v3(item.child_tag, item.parent_tag)
        if item != derived:
            raise DocxSemanticsV3Error("anchor-topology edge hash does not recompute")
        canonical.append(derived)
    entries = "".join(
        f'<edge child_tag="{item.child_tag}" parent_tag="{item.parent_tag}" sha256="{item.sha256}"/>'
        for item in sorted(canonical, key=lambda value: (value.child_tag, value.parent_tag))
    )
    root = (
        f'<documentAnchorTopologyMap xmlns="{ANCHOR_TOPOLOGY_MAP_NAMESPACE}" version="1">'
        f"{entries}</documentAnchorTopologyMap>"
    )
    declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    return f"{declaration}\n{root}\n".encode()


def parse_anchor_topology_map(root: Any) -> list[AnchorTopologyEdgeV3]:
    """Parse one closed canonical direct-parent topology map."""

    namespace = f"{{{ANCHOR_TOPOLOGY_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}documentAnchorTopologyMap"
        or tuple(root.attrib.items()) != (("version", "1"),)
        or root.text is not None
        or root.tail is not None
    ):
        raise DocxSemanticsV3Error("anchor-topology root is not closed and canonical")
    records: list[AnchorTopologyEdgeV3] = []
    for item in root:
        if (
            item.tag != f"{namespace}edge"
            or tuple(item.attrib) != ("child_tag", "parent_tag", "sha256")
            or item.text is not None
            or item.tail is not None
            or len(item) != 0
        ):
            raise DocxSemanticsV3Error("anchor-topology edge is not closed and canonical")
        edge = derive_anchor_topology_edge_v3(item.get("child_tag"), item.get("parent_tag"))
        if item.get("sha256") != edge.sha256:
            raise DocxSemanticsV3Error("anchor-topology edge hash does not recompute")
        records.append(edge)
    if not records:
        raise DocxSemanticsV3Error("anchor-topology map requires at least one edge")
    if records != sorted(records, key=lambda item: (item.child_tag, item.parent_tag)):
        raise DocxSemanticsV3Error("anchor-topology edges are not canonically ordered")
    _prove_topology_forest(records)
    return records


def prove_anchor_topology_document(
    body: Any,
    anchors: list[Any],
    records: list[AnchorTopologyEdgeV3],
    wrappers_by_tag: dict[str, Any],
    logical_by_tag: dict[str, tuple[Any, ...]],
    *,
    map_present: bool,
) -> None:
    """Bind the map exactly to nearest physical ordinary descendants."""

    from docx.oxml.ns import qn

    from docwen_core._docx_semantics_v3_ooxml import sdt_tag

    anchor_tags = {item.tag for item in anchors}
    if set(wrappers_by_tag) != anchor_tags or set(logical_by_tag) != anchor_tags:
        raise DocxSemanticsV3Error("anchor-topology endpoints differ from the target-map inventory")
    _prove_topology_forest(records, anchor_tags)
    physical: list[AnchorTopologyEdgeV3] = []
    for child_tag, wrapper in wrappers_by_tag.items():
        parent_tag = next(
            (
                tag
                for ancestor in wrapper.iterancestors(qn("w:sdt"))
                if (tag := sdt_tag(ancestor)) is not None and tag.startswith(ANCHOR_TAG_PREFIX)
            ),
            None,
        )
        if parent_tag is not None:
            physical.append(derive_anchor_topology_edge_v3(child_tag, parent_tag))
    expected = sorted(physical, key=lambda item: (item.child_tag, item.parent_tag))
    if bool(expected) != map_present:
        raise DocxSemanticsV3Error("anchor-topology map presence differs from physical nesting")
    if records != expected:
        raise DocxSemanticsV3Error("anchor-topology edges differ from nearest physical nesting")
    for edge in expected:
        _prove_flattened_child_range(logical_by_tag[edge.child_tag], logical_by_tag[edge.parent_tag])


def _prove_binding_pair(
    left: tuple[Any, int, int],
    right: tuple[Any, int, int],
    parent_by_tag: dict[str, str],
) -> None:
    left_binding, left_start, left_end = left
    right_binding, right_start, right_end = right
    left_tag = left_binding.identity.tag
    right_tag = right_binding.identity.tag
    if left_end < right_start or right_end < left_start:
        if _is_topology_ancestor(left_tag, right_tag, parent_by_tag) or _is_topology_ancestor(
            right_tag, left_tag, parent_by_tag
        ):
            raise DocxSemanticsV3Error("ordinary-anchor topology parent does not contain its child")
        return
    left_contains = left_start <= right_start and right_end <= left_end
    right_contains = right_start <= left_start and left_end <= right_end
    if not left_contains and not right_contains:
        raise DocxSemanticsV3Error("ordinary-anchor groups partially overlap")
    left_ancestor = _is_topology_ancestor(left_tag, right_tag, parent_by_tag)
    right_ancestor = _is_topology_ancestor(right_tag, left_tag, parent_by_tag)
    if left_start == right_start and left_end == right_end:
        if left_ancestor == right_ancestor:
            raise DocxSemanticsV3Error("ordinary-anchor groups have duplicate ownership without source topology")
    elif (left_contains and not left_ancestor) or (right_contains and not right_ancestor):
        raise DocxSemanticsV3Error("ordinary-anchor nested range lacks its authenticated source parent")


def _prove_topology_forest(
    records: list[AnchorTopologyEdgeV3],
    endpoint_tags: set[str] | None = None,
) -> None:
    parent_by_tag: dict[str, str] = {}
    for item in records:
        if endpoint_tags is not None and ({item.child_tag, item.parent_tag} - endpoint_tags):
            raise DocxSemanticsV3Error("anchor-topology edge has an unknown ordinary-anchor endpoint")
        if item.child_tag in parent_by_tag:
            raise DocxSemanticsV3Error("anchor-topology child has more than one parent")
        parent_by_tag[item.child_tag] = item.parent_tag
    for child in parent_by_tag:
        seen: set[str] = set()
        current = child
        while current in parent_by_tag:
            if current in seen:
                raise DocxSemanticsV3Error("anchor-topology graph contains a cycle")
            seen.add(current)
            current = parent_by_tag[current]


def _topology_depth(tag: str, parent_by_tag: dict[str, str]) -> int:
    depth = 0
    while tag in parent_by_tag:
        depth += 1
        tag = parent_by_tag[tag]
    return depth


def _is_topology_ancestor(ancestor: str, child: str, parent_by_tag: dict[str, str]) -> bool:
    while child in parent_by_tag:
        child = parent_by_tag[child]
        if child == ancestor:
            return True
    return False


def _prove_flattened_child_range(child: tuple[Any, ...], parent: tuple[Any, ...]) -> None:
    parent_positions = {id(item): index for index, item in enumerate(parent)}
    try:
        positions = [parent_positions[id(item)] for item in child]
    except KeyError as exc:
        raise DocxSemanticsV3Error("anchor-topology child is outside its flattened parent range") from exc
    if positions != list(range(positions[0], positions[0] + len(positions))):
        raise DocxSemanticsV3Error("anchor-topology child is not a contiguous flattened parent range")


def prove_owned_block_parent(child: Any, tag: str, body: Any, parent_sdt: Any | None, parent_tag: str | None) -> None:
    """Accept only direct body owners or exact anchor-under-anchor/target nesting."""

    from docx.oxml.ns import qn

    if parent_sdt is None:
        if child.getparent() is not body:
            raise DocxSemanticsV3Error("owned block SDT is not a direct main-body block")
        return
    content = child.getparent()
    if (
        not tag.startswith(ANCHOR_TAG_PREFIX)
        or parent_tag is None
        or not parent_tag.startswith((ANCHOR_TAG_PREFIX, TARGET_TAG_PREFIX))
        or content is None
        or content.tag != qn("w:sdtContent")
        or content.getparent() is not parent_sdt
    ):
        raise DocxSemanticsV3Error("owned block SDT has an invalid nesting owner")


def logical_group_elements(blocks: tuple[Any, ...]) -> tuple[Any, ...]:
    """Recursively flatten authenticated nested ordinary-anchor slots."""

    from docx.oxml.ns import qn

    from docwen_core._docx_semantics_v3_ooxml import sdt_tag

    output: list[Any] = []
    for block in blocks:
        if block.tag == qn("w:sdt") and (sdt_tag(block) or "").startswith(ANCHOR_TAG_PREFIX):
            content = block.find(qn("w:sdtContent"))
            if content is None or not list(content):
                raise DocxSemanticsV3Error("inner ordinary-anchor SDT has no logical content")
            output.extend(logical_group_elements(tuple(content)))
        else:
            output.append(block)
    return tuple(output)


def prove_ordinary_anchor_group(
    blocks: tuple[Any, ...],
    block_kind: str,
    *,
    allowed_heading_bookmark_name: str | None = None,
) -> tuple[Any, ...]:
    """Prove one anchor's recursively flattened logical kind and zero fields."""

    import re

    from docx.oxml.ns import qn

    logical = logical_group_elements(blocks)
    if not logical:
        raise DocxSemanticsV3Error("ordinary anchor has an empty logical block")
    if block_kind == "table":
        valid = len(logical) == 1 and logical[0].tag == qn("w:tbl")
    elif block_kind == "paragraph":
        valid = len(logical) == 1 and logical[0].tag == qn("w:p")
    elif block_kind in {"image", "equation", "code_block", "list_item"}:
        valid = bool(logical) and all(item.tag == qn("w:p") for item in logical)
    elif block_kind in {"list", "block_quote", "callout"}:
        valid = all(item.tag in {qn("w:p"), qn("w:tbl")} for item in logical)
    else:
        valid = False
    if not valid:
        raise DocxSemanticsV3Error("ordinary-anchor group has wrong block kind or cardinality")
    for item in logical:
        starts = list(item.iter(qn("w:bookmarkStart")))
        ends = list(item.iter(qn("w:bookmarkEnd")))
        has_disallowed_bookmark = bool(starts or ends) and (
            allowed_heading_bookmark_name is None
            or [start.get(qn("w:name")) for start in starts] != [allowed_heading_bookmark_name]
            or len(ends) != 1
        )
        if has_disallowed_bookmark or any(
            re.search(r"\b(?:SEQ|REF)\b", field.text or "", re.IGNORECASE) for field in item.iter(qn("w:instrText"))
        ):
            raise DocxSemanticsV3Error("ordinary anchor must not contain a bookmark, SEQ, or REF field")
    return logical


def prove_inline_occurrence_envelope(sdt: Any, tag: str) -> None:
    from docx.oxml.ns import qn

    children = list(sdt)
    if (
        sdt.attrib
        or sdt.text is not None
        or sdt.tail is not None
        or [item.tag for item in children] != [qn("w:sdtPr"), qn("w:sdtContent")]
    ):
        raise DocxSemanticsV3Error("reference-occurrence SDT envelope is not canonical")
    _prove_single_tag_properties(children[0], tag, context="reference-occurrence")
    content = children[1]
    if content.attrib or content.text is not None or content.tail is not None:
        raise DocxSemanticsV3Error("reference-occurrence SDT content is not canonical")


def prove_exact_field_run(
    run: Any,
    payload_tag: str,
    *,
    field_type: str | None = None,
    dirty: str | None = None,
    xml_space: str | None = None,
) -> None:
    from docx.oxml.ns import qn

    children = list(run)
    if run.attrib or run.xpath("text()") or run.tail is not None:
        raise DocxSemanticsV3Error("reference-occurrence REF run is not canonical")
    if len(children) != 1 or children[0].tag != payload_tag or len(children[0]) != 0:
        raise DocxSemanticsV3Error("reference-occurrence REF run payload is not canonical")
    payload = children[0]
    expected_attributes: tuple[tuple[str, str], ...]
    if field_type is not None:
        pairs = [(qn("w:fldCharType"), field_type)]
        if dirty is not None:
            pairs.append((qn("w:dirty"), dirty))
        expected_attributes = tuple(pairs)
    elif xml_space is not None:
        expected_attributes = (("{http://www.w3.org/XML/1998/namespace}space", xml_space),)
    else:
        expected_attributes = ()
    if (
        tuple(payload.attrib.items()) != expected_attributes
        or payload.tail is not None
        or (payload_tag == qn("w:fldChar") and payload.text is not None)
    ):
        raise DocxSemanticsV3Error("reference-occurrence REF payload attributes are not canonical")


def prove_alias_run(run: Any) -> None:
    from docx.oxml.ns import qn

    children = list(run)
    if run.attrib or run.xpath("text()") or run.tail is not None:
        raise DocxSemanticsV3Error("reference-occurrence Alias run is not canonical")
    if len(children) != 1 or children[0].tag != qn("w:t") or len(children[0]) != 0:
        raise DocxSemanticsV3Error("reference-occurrence Alias run payload is not canonical")
    text = children[0]
    if (
        tuple(text.attrib.items()) != (("{http://www.w3.org/XML/1998/namespace}space", "preserve"),)
        or text.tail is not None
    ):
        raise DocxSemanticsV3Error("reference-occurrence Alias spacing is not canonical")


def prove_soft_reference_envelope(sdt: Any, tag: str, expected_visible: str) -> None:
    from docx.oxml.ns import qn

    children = list(sdt)
    if (
        sdt.attrib
        or sdt.text is not None
        or sdt.tail is not None
        or [item.tag for item in children] != [qn("w:sdtPr"), qn("w:sdtContent")]
    ):
        raise DocxSemanticsV3Error("soft-reference SDT envelope is not canonical")
    _prove_single_tag_properties(children[0], tag, context="soft-reference")
    content = children[1]
    runs = list(content)
    if (
        content.attrib
        or content.text is not None
        or content.tail is not None
        or len(runs) != 1
        or runs[0].tag != qn("w:r")
        or runs[0].attrib
        or runs[0].xpath("text()")
        or runs[0].tail is not None
    ):
        raise DocxSemanticsV3Error("soft-reference SDT content is not canonical")
    texts = list(runs[0])
    expected_attributes = (
        (("{http://www.w3.org/XML/1998/namespace}space", "preserve"),)
        if expected_visible[:1].isspace() or expected_visible[-1:].isspace()
        else ()
    )
    if (
        len(texts) != 1
        or texts[0].tag != qn("w:t")
        or len(texts[0]) != 0
        or tuple(texts[0].attrib.items()) != expected_attributes
        or texts[0].tail is not None
        or (texts[0].text or "") != expected_visible
    ):
        raise DocxSemanticsV3Error("soft-reference cached text does not match authenticated metadata")


def prove_block_sdt_envelope(sdt: Any, tag: str) -> None:
    from docx.oxml.ns import qn

    children = list(sdt)
    if (
        sdt.attrib
        or sdt.text is not None
        or sdt.tail is not None
        or [item.tag for item in children] != [qn("w:sdtPr"), qn("w:sdtContent")]
    ):
        raise DocxSemanticsV3Error("owned block SDT envelope is not canonical")
    _prove_single_tag_properties(children[0], tag, context="owned block")
    content = children[1]
    if content.attrib or content.text is not None or content.tail is not None:
        raise DocxSemanticsV3Error("owned block SDT content is not canonical")


def prove_heading_paragraph(paragraph: Any) -> None:
    """Prove an addressable Heading remains a Heading1..Heading9 paragraph."""

    import re

    from docx.oxml.ns import qn

    style = paragraph.find(f"./{qn('w:pPr')}/{qn('w:pStyle')}")
    if style is None or re.fullmatch(r"Heading[1-9]", style.get(qn("w:val"), "")) is None:
        raise DocxSemanticsV3Error("Heading target must retain an exact Heading1..Heading9 style")


def prove_source_recovery_records(soft: list[Any], occurrences: list[Any]) -> None:
    """Prove one source identity and globally non-overlapping authored ranges."""

    records = sorted(
        [*soft, *occurrences],
        key=lambda item: (item.source_start, item.source_end, item.tag),
    )
    if len({item.source_sha256 for item in records}) > 1:
        raise DocxSemanticsV3Error("source-recovery maps do not share one source identity")
    for previous, current in pairwise(records):
        if current.source_start < previous.source_end:
            raise DocxSemanticsV3Error("source-recovery map ranges overlap")


def prove_source_recovery_physical_order(
    body: Any,
    soft: list[Any],
    occurrences: list[Any],
    fenced_sources: list[Any] | None = None,
) -> None:
    """Bind the merged source-range order to the physical inline SDT order."""

    from docx.oxml.ns import qn

    from docwen_core._docx_semantics_v3_fenced import FENCED_SOURCE_TAG_PREFIX
    from docwen_core._docx_semantics_v3_model import (
        REFERENCE_OCCURRENCE_TAG_PREFIX,
        SOFT_REFERENCE_TAG_PREFIX,
    )
    from docwen_core._docx_semantics_v3_ooxml import sdt_tag

    physical = [
        tag
        for sdt in body.iter(qn("w:sdt"))
        if (tag := sdt_tag(sdt)) is not None
        and tag.startswith(
            (
                SOFT_REFERENCE_TAG_PREFIX,
                REFERENCE_OCCURRENCE_TAG_PREFIX,
                FENCED_SOURCE_TAG_PREFIX,
            )
        )
    ]
    records: list[Any] = [*soft, *occurrences, *(fenced_sources or [])]
    expected = [
        item.tag
        for item in sorted(
            records,
            key=lambda item: (item.source_start, item.source_end, item.tag),
        )
    ]
    if physical != expected:
        raise DocxSemanticsV3Error("source-recovery SDT physical order differs from authenticated source order")


def _prove_single_tag_properties(properties: Any, tag: str, *, context: str) -> None:
    from docx.oxml.ns import qn

    children = list(properties)
    if (
        properties.attrib
        or properties.text is not None
        or properties.tail is not None
        or len(children) != 1
        or children[0].tag != qn("w:tag")
        or tuple(children[0].attrib.items()) != ((qn("w:val"), tag),)
        or children[0].text is not None
        or children[0].tail is not None
        or len(children[0]) != 0
    ):
        raise DocxSemanticsV3Error(f"{context} SDT properties are not canonical")
