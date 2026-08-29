"""Strict DOCX recovery and package authentication for semantics v3."""

from __future__ import annotations

import re
from collections.abc import Callable
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import docwen_core._docx_semantics_v3_fenced as fenced
from docwen_core._docx_caption_carrier import captionable_logical_elements
from docwen_core._docx_semantics_v3_fenced import FencedSourceIdentityV3
from docwen_core._docx_semantics_v3_fenced_map import parse_fenced_source_map
from docwen_core._docx_semantics_v3_model import (
    ANCHOR_TAG_PREFIX,
    ANCHOR_TOPOLOGY_MAP_NAMESPACE,
    CAPTION_STYLE_BINDING_MAP_NAMESPACE,
    REFERENCE_OCCURRENCE_MAP_NAMESPACE,
    REFERENCE_OCCURRENCE_TAG_PREFIX,
    SOFT_REFERENCE_MAP_NAMESPACE,
    SOFT_REFERENCE_TAG_PREFIX,
    TARGET_MAP_NAMESPACE,
    TARGET_TAG_PREFIX,
    AnchorIdentityV3,
    AnchorTopologyEdgeV3,
    CaptionStyleBindingV3,
    DocxSemanticsV3Error,
    OrdinaryAnchorGroupV3,
    RecoveredCaptionV3,
    ReferenceOccurrenceIdentityV3,
    SoftReferenceIdentityV3,
    SourceAnchorV3,
    TargetIdentityV3,
)
from docwen_core._docx_semantics_v3_ooxml import (
    recover_paragraph_children,
    sdt_tag,
    soft_reference_visible_text,
    visible_text,
)
from docwen_core._docx_semantics_v3_package import (
    parse_reference_occurrence_map,
    parse_semantic_map,
    parse_soft_reference_map,
    read_owned_map_parts,
    verify_custom_xml_support,
)
from docwen_core._docx_semantics_v3_styles import (
    caption_kind_for_paragraph_style,
    parse_caption_style_binding_map,
    prove_caption_paragraph_style,
    prove_caption_style_registry,
)
from docwen_core._docx_semantics_v3_topology import (
    logical_group_elements,
    parse_anchor_topology_map,
    prove_alias_run,
    prove_anchor_topology_document,
    prove_block_sdt_envelope,
    prove_exact_field_run,
    prove_heading_paragraph,
    prove_inline_occurrence_envelope,
    prove_ordinary_anchor_group,
    prove_owned_block_parent,
    prove_soft_reference_envelope,
    prove_source_recovery_physical_order,
    prove_source_recovery_records,
)
from docwen_core.docx_bookmarks import build_docx_bookmark_inventory, prove_bookmark_name

_TARGET_BOOKMARK_RE = re.compile(r"^DW_T_[0-9a-f]{35}$")


class DocxSemanticsV3Recovery:
    """Authenticated reader for packages emitted by the v3 session."""

    def __init__(
        self,
        *,
        block_elements: dict[Any, Any] | None = None,
        block_anchors: dict[Any, SourceAnchorV3] | None = None,
        ordinary_anchor_groups: dict[Any, tuple[OrdinaryAnchorGroupV3, ...]] | None = None,
        target_ids_by_bookmark: dict[str, str] | None = None,
        soft_tokens_by_tag: dict[str, str] | None = None,
        occurrence_tokens_by_tag: dict[str, str] | None = None,
        target_identities: tuple[TargetIdentityV3, ...] = (),
        anchor_identities: tuple[AnchorIdentityV3, ...] = (),
        anchor_topology_edges: tuple[AnchorTopologyEdgeV3, ...] = (),
        soft_reference_identities: tuple[SoftReferenceIdentityV3, ...] = (),
        reference_occurrence_identities: tuple[ReferenceOccurrenceIdentityV3, ...] = (),
        stable_reference_target_ids: tuple[str, ...] = (),
        recovered_captions: tuple[RecoveredCaptionV3, ...] = (),
        caption_style_bindings: tuple[CaptionStyleBindingV3, ...] = (),
        fenced_source_identities: tuple[FencedSourceIdentityV3, ...] = (),
        fenced_sources_by_paragraph: dict[Any, tuple[FencedSourceIdentityV3, str]] | None = None,
    ) -> None:
        self._block_elements = block_elements or {}
        self._block_anchors = block_anchors or {}
        self._ordinary_anchor_groups = ordinary_anchor_groups or {}
        self._target_ids_by_bookmark = target_ids_by_bookmark or {}
        self._soft_tokens_by_tag = soft_tokens_by_tag or {}
        self._occurrence_tokens_by_tag = occurrence_tokens_by_tag or {}
        self.target_identities = target_identities
        self.anchor_identities = anchor_identities
        self.anchor_topology_edges = anchor_topology_edges
        self.soft_reference_identities = soft_reference_identities
        self.reference_occurrence_identities = reference_occurrence_identities
        self.stable_reference_target_ids = stable_reference_target_ids
        self.recovered_captions = recovered_captions
        self.caption_style_bindings = caption_style_bindings
        self.fenced_source_identities = fenced_source_identities
        self._fenced_sources_by_paragraph = fenced_sources_by_paragraph or {}
        self.caption_signatures = tuple(
            (item.kind, item.source_id, item.title, item.cached_number) for item in recovered_captions
        )

    @classmethod
    def load(cls, path: str | Path, document: Any) -> DocxSemanticsV3Recovery:
        """Load zero or one strict map of each owned namespace."""

        caption_styles: tuple[CaptionStyleBindingV3, ...] = ()
        with ZipFile(Path(path)) as package:
            owned = read_owned_map_parts(package)
            for namespace, (item_number, _root) in owned.items():
                verify_custom_xml_support(package, item_number, namespace)
            if CAPTION_STYLE_BINDING_MAP_NAMESPACE in owned:
                _number, root = owned[CAPTION_STYLE_BINDING_MAP_NAMESPACE]
                caption_styles = parse_caption_style_binding_map(root)
                prove_caption_style_registry(package, caption_styles)

        targets: list[TargetIdentityV3] = []
        anchors: list[AnchorIdentityV3] = []
        soft: list[SoftReferenceIdentityV3] = []
        occurrences: list[ReferenceOccurrenceIdentityV3] = []
        fenced_sources: list[FencedSourceIdentityV3] = []
        anchor_topology: list[AnchorTopologyEdgeV3] = []
        if TARGET_MAP_NAMESPACE in owned:
            _number, root = owned[TARGET_MAP_NAMESPACE]
            targets, anchors = parse_semantic_map(root)
        if ANCHOR_TOPOLOGY_MAP_NAMESPACE in owned:
            _number, root = owned[ANCHOR_TOPOLOGY_MAP_NAMESPACE]
            anchor_topology = parse_anchor_topology_map(root)
        if SOFT_REFERENCE_MAP_NAMESPACE in owned:
            _number, root = owned[SOFT_REFERENCE_MAP_NAMESPACE]
            soft = parse_soft_reference_map(root)
        if REFERENCE_OCCURRENCE_MAP_NAMESPACE in owned:
            _number, root = owned[REFERENCE_OCCURRENCE_MAP_NAMESPACE]
            occurrences = parse_reference_occurrence_map(root)
        if fenced.FENCED_SOURCE_MAP_NAMESPACE in owned:
            _number, root = owned[fenced.FENCED_SOURCE_MAP_NAMESPACE]
            fenced_sources = parse_fenced_source_map(root)
        return cls._bind_document_evidence(
            document,
            targets,
            anchors,
            soft,
            occurrences,
            caption_styles,
            fenced_sources,
            anchor_topology,
            topology_map_present=ANCHOR_TOPOLOGY_MAP_NAMESPACE in owned,
        )

    @classmethod
    def _bind_document_evidence(
        cls,
        document: Any,
        targets: list[TargetIdentityV3],
        anchors: list[AnchorIdentityV3],
        soft_references: list[SoftReferenceIdentityV3],
        reference_occurrences: list[ReferenceOccurrenceIdentityV3],
        caption_style_bindings: tuple[CaptionStyleBindingV3, ...],
        fenced_sources: list[FencedSourceIdentityV3],
        anchor_topology: list[AnchorTopologyEdgeV3],
        *,
        topology_map_present: bool,
        caption_parser: Callable[..., tuple[str, str]] | None = None,
    ) -> DocxSemanticsV3Recovery:
        from docx.oxml.ns import qn

        target_by_tag = {item.tag: item for item in targets}
        anchor_by_tag = {item.tag: item for item in anchors}
        block_elements: dict[Any, tuple[Any, ...]] = {}
        block_anchors: dict[Any, SourceAnchorV3] = {}
        seen_targets: set[str] = set()
        seen_anchors: set[str] = set()
        ordinary_anchor_groups: dict[Any, list[OrdinaryAnchorGroupV3]] = {}
        anchor_group_records: list[tuple[Any, SourceAnchorV3, tuple[Any, ...]]] = []
        anchor_wrappers_by_tag: dict[str, Any] = {}
        anchor_logical_by_tag: dict[str, tuple[Any, ...]] = {}
        body = document.element.body
        owned_block_sdts = [
            child
            for child in body.iter(qn("w:sdt"))
            if (sdt_tag(child) or "").startswith((TARGET_TAG_PREFIX, ANCHOR_TAG_PREFIX))
        ]
        for child in owned_block_sdts:
            tag = sdt_tag(child)
            assert tag is not None
            parent_sdt = next(
                (ancestor for ancestor in child.iterancestors(qn("w:sdt"))),
                None,
            )
            prove_owned_block_parent(
                child, tag, body, parent_sdt, sdt_tag(parent_sdt) if parent_sdt is not None else None
            )
            prove_block_sdt_envelope(child, tag)
            content = child.find(qn("w:sdtContent"))
            if content is None:
                raise DocxSemanticsV3Error("owned block SDT is missing sdtContent")
            blocks = tuple(content)
            if not blocks:
                raise DocxSemanticsV3Error("owned block SDT has no logical content")
            block_elements[child] = blocks
            if tag in target_by_tag:
                if tag in seen_targets:
                    raise DocxSemanticsV3Error("duplicate target outer SDT")
                target = target_by_tag[tag]
                seen_targets.add(tag)
                if target.kind == "heading":
                    if len(blocks) != 1:
                        raise DocxSemanticsV3Error("Heading target SDT must contain exactly one block")
                    direct_tag = sdt_tag(blocks[0]) if blocks[0].tag == qn("w:sdt") else None
                    if blocks[0].tag != qn("w:p") and direct_tag not in anchor_by_tag:
                        raise DocxSemanticsV3Error(
                            "Heading target SDT inner block is not an authenticated ordinary anchor"
                        )
                    logical = logical_group_elements(blocks)
                    if len(logical) != 1 or logical[0].tag != qn("w:p"):
                        raise DocxSemanticsV3Error("Heading target SDT must resolve to exactly one paragraph")
                    prove_heading_paragraph(logical[0])
                    block_anchors[logical[0]] = SourceAnchorV3("semantic_target", target.source_id, target.kind)
                else:
                    caption, _object_elements = _prove_caption_target_group(
                        blocks,
                        target,
                        caption_style_bindings,
                        caption_parser=caption_parser,
                    )
                    block_anchors[caption] = SourceAnchorV3("semantic_target", target.source_id, target.kind)
            elif tag in anchor_by_tag:
                if tag in seen_anchors:
                    raise DocxSemanticsV3Error("duplicate ordinary-anchor outer SDT")
                anchor = anchor_by_tag[tag]
                seen_anchors.add(tag)
                heading_bookmark = next(
                    (
                        owner.bookmark_name
                        for ancestor in child.iterancestors(qn("w:sdt"))
                        if (owner := target_by_tag.get(sdt_tag(ancestor) or "")) is not None and owner.kind == "heading"
                    ),
                    None,
                )
                logical = prove_ordinary_anchor_group(
                    blocks,
                    anchor.block_kind,
                    allowed_heading_bookmark_name=heading_bookmark,
                )
                source_anchor = SourceAnchorV3("ordinary_anchor", anchor.source_id, anchor.block_kind)
                anchor_group_records.append((child, source_anchor, logical))
                anchor_wrappers_by_tag[tag] = child
                anchor_logical_by_tag[tag] = logical
            else:
                raise DocxSemanticsV3Error("owned block SDT has no authenticated map record")
        if seen_targets != set(target_by_tag) or seen_anchors != set(anchor_by_tag):
            raise DocxSemanticsV3Error("v3 map record is missing its exact block SDT")
        prove_anchor_topology_document(
            body,
            anchors,
            anchor_topology,
            anchor_wrappers_by_tag,
            anchor_logical_by_tag,
            map_present=topology_map_present,
        )
        for _wrapper, source_anchor, logical in reversed(anchor_group_records):
            for index, block in enumerate(logical):
                group = OrdinaryAnchorGroupV3(source_anchor, logical, index)
                ordinary_anchor_groups.setdefault(block, []).append(group)
                block_anchors.setdefault(block, source_anchor)

        target_ids_by_bookmark = _prove_target_bookmarks(document, targets, block_elements, block_anchors)
        recovered_captions = _recover_captions(
            body,
            targets,
            block_elements,
            caption_style_bindings,
            caption_parser=caption_parser,
        )
        prove_source_recovery_records(soft_references, [*reference_occurrences, *fenced_sources])
        fenced_sources_by_paragraph = fenced.bind_fenced_source_document_v3(body, fenced_sources)
        prove_source_recovery_physical_order(body, soft_references, reference_occurrences, fenced_sources)
        soft_tokens_by_tag = _prove_soft_references(body, soft_references)
        occurrence_tokens_by_tag, occurrence_target_ids = _prove_reference_occurrences(
            body,
            reference_occurrences,
            targets,
        )
        stable_target_ids: list[str] = []
        for paragraph in body.iter(qn("w:p")):
            recover_paragraph_children(
                list(paragraph),
                target_ids_by_bookmark=target_ids_by_bookmark,
                soft_tokens_by_tag=soft_tokens_by_tag,
                occurrence_tokens_by_tag=occurrence_tokens_by_tag,
                stable_reference_target_ids=stable_target_ids,
            )
        if stable_target_ids != occurrence_target_ids:
            raise DocxSemanticsV3Error("reference-occurrence and recovered REF inventories differ")
        return cls(
            block_elements=block_elements,
            block_anchors=block_anchors,
            ordinary_anchor_groups={item: tuple(groups) for item, groups in ordinary_anchor_groups.items()},
            target_ids_by_bookmark=target_ids_by_bookmark,
            soft_tokens_by_tag=soft_tokens_by_tag,
            occurrence_tokens_by_tag=occurrence_tokens_by_tag,
            target_identities=tuple(sorted(targets, key=lambda item: (item.kind, item.source_id))),
            anchor_identities=tuple(sorted(anchors, key=lambda item: (item.block_kind, item.source_id))),
            anchor_topology_edges=tuple(anchor_topology),
            soft_reference_identities=tuple(
                sorted(soft_references, key=lambda item: (item.source_start, item.source_end, item.tag))
            ),
            reference_occurrence_identities=tuple(
                sorted(
                    reference_occurrences,
                    key=lambda item: (item.source_start, item.source_end, item.tag),
                )
            ),
            stable_reference_target_ids=tuple(stable_target_ids),
            recovered_captions=tuple(recovered_captions),
            caption_style_bindings=caption_style_bindings,
            fenced_source_identities=tuple(fenced_sources),
            fenced_sources_by_paragraph=fenced_sources_by_paragraph,
        )

    def logical_body_elements(self, document: Any) -> list[Any]:
        """Flatten only authenticated owned block SDTs for the legacy walker."""

        output: list[Any] = []
        for child in list(document.element.body):
            blocks = self._block_elements.get(child)
            if blocks is None:
                output.append(child)
            else:
                output.extend(logical_group_elements(blocks))
        return output

    def source_anchor(self, paragraph_element: Any) -> SourceAnchorV3 | None:
        return self._block_anchors.get(paragraph_element)

    def ordinary_anchor_group(self, element: Any) -> OrdinaryAnchorGroupV3 | None:
        """Return the innermost authenticated ordinary-anchor owner, if any."""

        groups = self.ordinary_anchor_groups(element)
        return groups[0] if groups else None

    def ordinary_anchor_groups(self, element: Any) -> tuple[OrdinaryAnchorGroupV3, ...]:
        """Return every authenticated owner in deterministic inner-to-outer order."""

        return self._ordinary_anchor_groups.get(element, ())

    def caption_for_object(self, element: Any) -> RecoveredCaptionV3 | None:
        return next(
            (item for item in self.recovered_captions if element in item.object_elements),
            None,
        )

    def is_caption_element(self, element: Any) -> bool:
        return any(item.caption_element is element for item in self.recovered_captions)

    def render_paragraph_text(self, paragraph_element: Any) -> str | None:
        """Recover exact authored v3 reference tokens from one paragraph."""

        text, semantic = recover_paragraph_children(
            list(paragraph_element),
            target_ids_by_bookmark=self._target_ids_by_bookmark,
            soft_tokens_by_tag=self._soft_tokens_by_tag,
            occurrence_tokens_by_tag=self._occurrence_tokens_by_tag,
        )
        return text if semantic else None

    def render_fenced_source(self, paragraph_element: Any) -> str | None:
        """Recover an exact authored fence only from authenticated carrier evidence."""

        item = self._fenced_sources_by_paragraph.get(paragraph_element)
        return None if item is None else item[1]


def _prove_caption_target_group(
    blocks: tuple[Any, ...],
    target: TargetIdentityV3,
    caption_style_bindings: tuple[CaptionStyleBindingV3, ...],
    *,
    caption_parser: Callable[..., tuple[str, str]] | None = None,
) -> tuple[Any, tuple[Any, ...]]:
    from docx.oxml.ns import qn

    if len(blocks) != 2:
        raise DocxSemanticsV3Error("caption target SDT must contain one caption and one logical object")
    if target.kind == "figure":
        object_slot, caption = blocks
    else:
        caption, object_slot = blocks
    if caption.tag != qn("w:p"):
        raise DocxSemanticsV3Error("caption target has no caption paragraph")
    if not caption_style_bindings:
        raise DocxSemanticsV3Error("caption target lacks its authenticated caption-style binding map")
    prove_caption_paragraph_style(caption, target.kind, caption_style_bindings)
    object_elements = logical_group_elements((object_slot,))
    _prove_caption_object_kind(target.kind, object_elements)
    parser = _parse_v3_caption if caption_parser is None else caption_parser
    parser(caption, target.kind, required_bookmark=target.bookmark_name)
    return caption, object_elements


def _prove_caption_object_kind(kind: str, elements: tuple[Any, ...]) -> None:
    if kind not in {"figure", "table", "equation", "code_block"}:
        raise DocxSemanticsV3Error("caption target has an unsupported semantic kind")
    captionable_logical_elements(elements)


def _parse_v3_caption(
    paragraph: Any,
    kind: str,
    *,
    required_bookmark: str | None,
) -> tuple[str, str]:
    from docx.oxml.ns import qn

    counter = {"figure": "Figure", "table": "Table", "equation": "Equation", "code_block": "Code"}[kind]
    instructions = [item.text or "" for item in paragraph.iter(qn("w:instrText"))]
    expected = f" SEQ {counter} \\* ARABIC "
    if instructions != [expected]:
        raise DocxSemanticsV3Error("caption must contain exactly one matching SEQ field")
    field_chars = [item.get(qn("w:fldCharType")) for item in paragraph.iter(qn("w:fldChar"))]
    if field_chars != ["begin", "separate", "end"]:
        raise DocxSemanticsV3Error("caption SEQ field skeleton is not canonical")
    if required_bookmark is None:
        if paragraph.find(f".//{qn('w:bookmarkStart')}") is not None:
            raise DocxSemanticsV3Error("ID-less caption must not contain a bookmark")
    else:
        names = [item.get(qn("w:name")) for item in paragraph.iter(qn("w:bookmarkStart"))]
        if names != [required_bookmark]:
            raise DocxSemanticsV3Error("caption target bookmark does not enclose its SEQ field")
    texts = [item.text or "" for item in paragraph.iter(qn("w:t"))]
    if len(texts) < 2 or texts[0] != f"{counter} ":
        raise DocxSemanticsV3Error("caption visible label does not match its kind")
    cached_number = texts[1]
    suffix = "".join(texts[2:])
    title = suffix[2:] if suffix.startswith(": ") else ""
    if suffix and not suffix.startswith(": "):
        raise DocxSemanticsV3Error("caption visible title separator is not canonical")
    return title, cached_number


def _recover_captions(
    body: Any,
    targets: list[TargetIdentityV3],
    block_elements: dict[Any, tuple[Any, ...]],
    caption_style_bindings: tuple[CaptionStyleBindingV3, ...],
    *,
    caption_parser: Callable[..., tuple[str, str]] | None = None,
) -> list[RecoveredCaptionV3]:
    from docx.oxml.ns import qn

    recovered: list[RecoveredCaptionV3] = []
    claimed: set[int] = set()
    target_by_tag = {item.tag: item for item in targets if item.kind != "heading"}
    for outer, blocks in block_elements.items():
        tag = sdt_tag(outer)
        target = target_by_tag.get(tag or "")
        if target is None:
            continue
        caption, object_elements = _prove_caption_target_group(
            blocks,
            target,
            caption_style_bindings,
            caption_parser=caption_parser,
        )
        parser = _parse_v3_caption if caption_parser is None else caption_parser
        title, number = parser(
            caption,
            target.kind,
            required_bookmark=target.bookmark_name,
        )
        recovered.append(
            RecoveredCaptionV3(
                kind=target.kind,  # type: ignore[arg-type]
                source_id=target.source_id,
                title=title,
                cached_number=number,
                caption_element=caption,
                object_elements=object_elements,
            )
        )
        claimed.update((id(caption), *(id(item) for item in object_elements)))

    if not caption_style_bindings:
        return recovered
    children = list(body)
    for index, child in enumerate(children):
        if id(child) in claimed or child.tag != qn("w:p"):
            continue
        kind = caption_kind_for_paragraph_style(child, caption_style_bindings)
        if kind is None:
            continue
        object_index = index - 1 if kind == "figure" else index + 1
        if not 0 <= object_index < len(children):
            raise DocxSemanticsV3Error("ID-less caption has no directly adjacent object")
        object_slot = children[object_index]
        if id(object_slot) in claimed:
            raise DocxSemanticsV3Error("caption/object participates in multiple pairing claims")
        object_elements = logical_group_elements((object_slot,))
        _prove_caption_object_kind(kind, object_elements)
        parser = _parse_v3_caption if caption_parser is None else caption_parser
        title, number = parser(child, kind, required_bookmark=None)
        recovered.append(
            RecoveredCaptionV3(
                kind=kind,  # type: ignore[arg-type]
                source_id=None,
                title=title,
                cached_number=number,
                caption_element=child,
                object_elements=object_elements,
            )
        )
        claimed.update((id(child), *(id(item) for item in object_elements)))
    return recovered


def _prove_target_bookmarks(
    document: Any,
    targets: list[TargetIdentityV3],
    block_elements: dict[Any, tuple[Any, ...]],
    block_anchors: dict[Any, SourceAnchorV3],
) -> dict[str, str]:
    from docx.oxml.ns import qn

    inventory = build_docx_bookmark_inventory(document)
    target_ids: dict[str, str] = {}
    for target in targets:
        if not prove_bookmark_name(inventory, target.bookmark_name).valid:
            raise DocxSemanticsV3Error("target bookmark is not unique, balanced, and ordered")
        paragraph = next(
            element
            for group in block_elements.values()
            for element in logical_group_elements(group)
            if (anchor := block_anchors.get(element)) is not None
            and anchor.owner_kind == "semantic_target"
            and anchor.source_id == target.source_id
        )
        names = {item.get(qn("w:name")) for item in paragraph.iter(qn("w:bookmarkStart"))}
        if target.bookmark_name not in names:
            raise DocxSemanticsV3Error("target bookmark is outside its authenticated outer SDT")
        target_ids[target.bookmark_name] = target.source_id
    owned = {
        item.name
        for item in inventory.starts
        if item.name is not None and _TARGET_BOOKMARK_RE.fullmatch(item.name) is not None
    }
    if owned != set(target_ids):
        raise DocxSemanticsV3Error("target bookmark and target-map inventories differ")
    return target_ids


def _prove_soft_references(
    body: Any,
    records: list[SoftReferenceIdentityV3],
) -> dict[str, str]:
    from docx.oxml.ns import qn

    by_tag = {item.tag: item for item in records}
    seen: set[str] = set()
    physical_tags: list[str] = []
    for sdt in body.iter(qn("w:sdt")):
        tag = sdt_tag(sdt)
        if not tag or not tag.startswith(SOFT_REFERENCE_TAG_PREFIX):
            continue
        record = by_tag.get(tag)
        if record is None or tag in seen:
            raise DocxSemanticsV3Error("soft-reference SDT is missing, duplicated, or unmapped")
        if sdt.getparent() is None or sdt.getparent().tag != qn("w:p"):
            raise DocxSemanticsV3Error("soft-reference SDT must be a direct paragraph child")
        physical_tags.append(tag)
        prove_soft_reference_envelope(
            sdt,
            tag,
            soft_reference_visible_text(record.authored_token, record.cached_number),
        )
        seen.add(tag)
    if seen != set(by_tag):
        raise DocxSemanticsV3Error("soft-reference map record is missing its exact inline SDT")
    expected_tags = [
        item.tag for item in sorted(records, key=lambda item: (item.source_start, item.source_end, item.tag))
    ]
    if physical_tags != expected_tags:
        raise DocxSemanticsV3Error("soft-reference physical order differs from authenticated source order")
    return {item.tag: item.authored_token for item in records}


def _prove_reference_occurrences(
    body: Any,
    records: list[ReferenceOccurrenceIdentityV3],
    targets: list[TargetIdentityV3],
) -> tuple[dict[str, str], list[str]]:
    from docx.oxml.ns import qn

    target_names = {item.bookmark_name for item in targets}
    by_tag = {item.tag: item for item in records}
    seen: set[str] = set()
    target_ids: list[str] = []
    physical_tags: list[str] = []
    for sdt in body.iter(qn("w:sdt")):
        tag = sdt_tag(sdt)
        if not tag or not tag.startswith(REFERENCE_OCCURRENCE_TAG_PREFIX):
            continue
        record = by_tag.get(tag)
        if record is None or tag in seen:
            raise DocxSemanticsV3Error("reference-occurrence SDT is missing, duplicated, or unmapped")
        if sdt.getparent() is None or sdt.getparent().tag != qn("w:p"):
            raise DocxSemanticsV3Error("reference-occurrence SDT must be a direct paragraph child")
        physical_tags.append(tag)
        if record.resolved_bookmark_name not in target_names:
            raise DocxSemanticsV3Error("reference-occurrence points outside the authenticated target map")
        content = sdt.find(qn("w:sdtContent"))
        if content is None:
            raise DocxSemanticsV3Error("reference-occurrence SDT has no content")
        prove_inline_occurrence_envelope(sdt, tag)
        instruction = "".join(item.text or "" for item in content.iter(qn("w:instrText")))
        expected_heading = next(
            item.kind == "heading" for item in targets if item.bookmark_name == record.resolved_bookmark_name
        )
        target_ids.append(
            next(item.source_id for item in targets if item.bookmark_name == record.resolved_bookmark_name)
        )
        expected = (
            f" REF {record.resolved_bookmark_name} \\n \\h "
            if expected_heading
            else f" REF {record.resolved_bookmark_name} \\h "
        )
        if instruction != expected:
            raise DocxSemanticsV3Error("reference-occurrence REF instruction does not match authenticated metadata")
        _prove_reference_occurrence_field_skeleton(content, record)
        expected_visible = _reference_occurrence_visible(record)
        if visible_text(content) != expected_visible:
            raise DocxSemanticsV3Error("reference-occurrence visible result does not match authenticated metadata")
        seen.add(tag)
    if seen != set(by_tag):
        raise DocxSemanticsV3Error("reference-occurrence map record is missing its exact inline SDT")
    expected_tags = [
        item.tag for item in sorted(records, key=lambda item: (item.source_start, item.source_end, item.tag))
    ]
    if physical_tags != expected_tags:
        raise DocxSemanticsV3Error("reference-occurrence physical order differs from authenticated source order")
    return {item.tag: item.authored_token for item in records}, target_ids


def _reference_occurrence_visible(record: ReferenceOccurrenceIdentityV3) -> str:
    body = record.authored_token[3:-2]
    _selector, separator, alias = body.partition("|")
    return record.cached_number if not separator else f"{record.cached_number} {alias}"


def _prove_reference_occurrence_field_skeleton(
    content: Any,
    record: ReferenceOccurrenceIdentityV3,
) -> None:
    from docx.oxml.ns import qn

    children = list(content)
    body = record.authored_token[3:-2]
    _selector, separator, alias = body.partition("|")
    expected_length = 6 if separator else 5
    if len(children) != expected_length or any(item.tag != qn("w:r") for item in children):
        raise DocxSemanticsV3Error("reference-occurrence must contain canonical REF runs and optional Alias only")
    prove_exact_field_run(children[0], qn("w:fldChar"), field_type="begin", dirty="true")
    prove_exact_field_run(children[1], qn("w:instrText"), xml_space="preserve")
    prove_exact_field_run(children[2], qn("w:fldChar"), field_type="separate")
    prove_exact_field_run(children[3], qn("w:t"))
    prove_exact_field_run(children[4], qn("w:fldChar"), field_type="end")
    instruction_texts = list(children[1].iter(qn("w:instrText")))
    result_texts = list(children[3].iter(qn("w:t")))
    if len(instruction_texts) != 1 or len(result_texts) != 1 or (result_texts[0].text or "") != record.cached_number:
        raise DocxSemanticsV3Error("reference-occurrence cached number is not the exact REF result")
    if separator:
        prove_alias_run(children[5])
        alias_texts = list(children[5].iter(qn("w:t")))
        if len(alias_texts) != 1 or (alias_texts[0].text or "") != f" {alias}":
            raise DocxSemanticsV3Error("reference-occurrence Alias is not the exact authored suffix")
