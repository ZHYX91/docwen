"""Request-owned DOCX session for a validated v4 resolved-numbering port."""

from __future__ import annotations

import hashlib
import os
import tempfile
from collections.abc import Callable
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docwen_core._docx_recovery_map import (
    ResolvedV4RecoveryInput,
    ResolvedV4RecoveryMap,
    build_recovery_map,
    compute_physical_projection,
    inject_recovery_map,
)
from docwen_core._docx_resolved_inline import (
    ResolvedInlineCarrier,
    ResolvedInlineFragment,
    ResolvedInlineProjectionError,
    capture_resolved_inline_projection,
)
from docwen_core._docx_resolved_numbering_proof import (
    ResolvedNumberingDocxError,
    ResolvedNumberingProofMixin,
)
from docwen_core._docx_resolved_source_carriers import (
    ResolvedOrdinaryAnchorBindingV4,
    admit_resolved_ordinary_anchor_v4,
    order_resolved_block_carriers_v4,
    physical_resolved_block_tags_v4,
    resolved_anchor_group_v4,
    resolved_semantic_group_v4,
    validate_fenced_anchor_ranges_v4,
)
from docwen_core._docx_semantics_v3_fenced import (
    FENCED_SOURCE_MAP_NAMESPACE,
    FencedSourceBindingV3,
    FencedSourceIdentityV3,
    prove_fenced_source_paragraph,
    reconstruct_fenced_source_v3,
    wrap_fenced_paragraph_payload,
    write_canonical_fenced_body_v3,
)
from docwen_core._docx_semantics_v3_fenced_map import (
    fenced_source_identity_from_mapping_v3,
    fenced_source_map_xml,
    fenced_source_mapping_v3,
)
from docwen_core._docx_semantics_v3_model import (
    ANCHOR_TOPOLOGY_MAP_NAMESPACE,
    CAPTION_STYLE_BINDING_MAP_NAMESPACE,
    REFERENCE_OCCURRENCE_MAP_NAMESPACE,
    SOFT_REFERENCE_MAP_NAMESPACE,
    TARGET_MAP_NAMESPACE,
    AnchorTopologyEdgeV3,
    CaptionStyleBindingV3,
    DocxSemanticsV3Error,
    ReferenceOccurrenceIdentityV3,
    SoftReferenceIdentityV3,
    TargetBindingV3,
    TargetIdentityV3,
    derive_reference_occurrence_identity_v3,
    derive_soft_reference_identity_v3,
    derive_target_identity_v3,
)
from docwen_core._docx_semantics_v3_ooxml import (
    inline_sdt,
    soft_reference_visible_text,
    visible_text,
    wrap_direct_body_group,
    wrap_paragraph_content_with_bookmark,
)
from docwen_core._docx_semantics_v3_package import (
    inject_custom_xml_parts,
    reference_occurrence_map_xml,
    semantic_map_xml,
    soft_reference_map_xml,
)
from docwen_core._docx_semantics_v3_styles import (
    caption_style_binding_map_xml,
    prove_caption_paragraph_style,
    validate_caption_style_bindings,
)
from docwen_core._docx_semantics_v3_topology import anchor_topology_map_xml
from docwen_core.docx_bookmarks import (
    BOOKMARK_ID_MAX,
    bookmark_id_key,
    build_docx_bookmark_inventory,
)
from docwen_core.docx_citation_ooxml import (
    CITATION_ITEM_MAP_NAMESPACE,
    CITATION_OCCURRENCE_MAP_NAMESPACE,
    ResolvedCitationProjection,
    build_resolved_citation_projection,
    citation_item_map_xml,
    citation_occurrence_map_xml,
    citation_occurrence_sdt,
    preflight_citation_document,
)
from docwen_core.docx_numbering_occurrence import (
    NUMBERING_OCCURRENCE_MAP_NAMESPACE,
    NumberingOccurrenceIdentity,
    derive_numbering_occurrence,
    numbering_occurrence_map_xml,
    wrap_numbering_occurrence,
)
from docwen_core.docx_numbering_ooxml import (
    HeadingNumberingProjection,
    apply_heading_numbering,
    create_heading_numbering_projection,
    existing_numbering_ids,
    inline_reference_sdt,
    materialize_caption_number,
    write_heading_numbering_projection,
)
from docwen_core.models._resolved_numbering_semantics import (
    validate_document,
    validate_plan,
    validate_port,
)
from docwen_core.models.resolved_numbering import (
    NumberingTarget,
    ResolvedDocumentTarget,
    ResolvedNumberingPort,
    ResolvedNumberingPortError,
)


@dataclass(frozen=True, slots=True)
class _DisabledOccurrenceBinding:
    identity: NumberingOccurrenceIdentity
    caption_element: Any
    object_elements: tuple[Any, ...]


@dataclass(frozen=True, slots=True)
class _CaptionPlanBinding:
    document_target: ResolvedDocumentTarget
    plan_target: NumberingTarget
    object_count: int
    object_snapshots: tuple[Any, ...]
    payload_fragments: tuple[ResolvedInlineFragment, ...] | None


class ResolvedNumberingDocxSession(ResolvedNumberingProofMixin):
    """Bind a typed numbering snapshot to one python-docx request.

    Callers render authored text and structural objects, then bind each target
    and reference by its authenticated source range.  The session never parses
    Markdown, computes counters, or interprets a source prefix.
    """

    def __init__(
        self,
        document: Any,
        port: ResolvedNumberingPort,
        *,
        heading_style_ids: dict[int, str],
        heading_style_names: dict[str, str],
        caption_style_bindings: tuple[CaptionStyleBindingV3, ...],
        recovery_input: ResolvedV4RecoveryInput | None = None,
    ) -> None:
        try:
            validate_document(port.document, port.source_sha256, "docwen.resolved_document.invalid")
            validate_plan(port.plan)
            validate_port(port.document_envelope, port.plan_envelope)
        except ResolvedNumberingPortError as exc:
            raise ResolvedNumberingDocxError("resolved-numbering port failed runtime revalidation") from exc
        self._document = document
        self._port = port
        self._heading_style_ids = dict(heading_style_ids)
        self._heading_style_names = dict(heading_style_names)
        self._caption_style_bindings = validate_caption_style_bindings(caption_style_bindings)
        existing_abstract, existing_num = existing_numbering_ids(document)
        self._projection = create_heading_numbering_projection(
            port.plan,
            heading_style_ids=self._heading_style_ids,
            existing_abstract_ids=existing_abstract,
            existing_num_ids=existing_num,
        )
        self._document_targets = {item.occurrence_key: item for item in port.document.targets}
        self._plan_targets = {item.occurrence_key: item for item in port.plan.targets}
        if self._document_targets.keys() != self._plan_targets.keys():
            raise ResolvedNumberingDocxError("resolved document and plan target inventories differ")
        self._references = {(item.source_start, item.source_end): item for item in port.document.references}
        self._citations = {(item.source_start, item.source_end): item for item in port.document.citations}
        self._citation_projection = build_resolved_citation_projection(port.source_sha256, port.document.citations)
        if self._citation_projection is not None:
            preflight_citation_document(document, self._citation_projection)
        self._citation_occurrences = (
            {}
            if self._citation_projection is None
            else {item.occurrence_key: item for item in self._citation_projection.occurrence_map.occurrences}
        )
        self._bound_targets: set[tuple[int, int, str]] = set()
        self._bound_references: set[tuple[int, int]] = set()
        self._bound_citations: set[tuple[int, int]] = set()
        self._reference_paragraphs: dict[tuple[int, int], Any] = {}
        self._citation_paragraphs: dict[tuple[int, int], Any] = {}
        self._target_bindings: list[TargetBindingV3] = []
        self._ordinary_anchors: list[ResolvedOrdinaryAnchorBindingV4] = []
        self._anchor_topology_edges: tuple[AnchorTopologyEdgeV3, ...] = ()
        self._fenced_sources: list[FencedSourceBindingV3] = []
        self._block_parent_by_tag: dict[str, str | None] = {}
        self._block_physical_tags: tuple[str, ...] = ()
        self._caption_target_keys: list[tuple[int, int, str]] = []
        self._caption_plan_bindings: list[_CaptionPlanBinding] = []
        self._heading_rendered_texts: dict[tuple[int, int, str], str] = {}
        self._heading_payload_snapshots: dict[tuple[int, int, str], tuple[Any, ...]] = {}
        self._source_projected_heading_keys: set[tuple[int, int, str]] = set()
        self._occurrence_bindings: list[_DisabledOccurrenceBinding] = []
        self._reference_occurrences: list[ReferenceOccurrenceIdentityV3] = []
        self._soft_references: list[SoftReferenceIdentityV3] = []
        self._stable_reference_target_ids: list[str] = []
        self._claimed_ids: set[str] = set()
        self._recovery_input = recovery_input
        self._recovery_map: ResolvedV4RecoveryMap | None = None
        self._finalized = False

    @property
    def projection(self) -> HeadingNumberingProjection:
        return self._projection

    @property
    def port(self) -> ResolvedNumberingPort:
        return self._port

    @property
    def citation_projection(self) -> ResolvedCitationProjection | None:
        return self._citation_projection

    @property
    def recovery_map(self) -> ResolvedV4RecoveryMap | None:
        return self._recovery_map

    def bind_ordinary_anchor(
        self,
        elements: tuple[Any, ...],
        anchor: dict[str, Any],
        *,
        direct_parent_source_id: str | None = None,
    ) -> None:
        """Bind one frozen source-only ordinary anchor to this resolved session."""

        if self._finalized:
            raise ResolvedNumberingDocxError("resolved-numbering session is finalized")
        try:
            binding = admit_resolved_ordinary_anchor_v4(
                self._port.document.authored_markdown,
                elements,
                anchor,
                direct_parent_source_id=direct_parent_source_id,
            )
        except DocxSemanticsV3Error as exc:
            raise ResolvedNumberingDocxError("resolved ordinary-anchor carrier is invalid") from exc
        self._claim_source_id(binding.identity.source_id)
        self._ordinary_anchors.append(binding)

    def bind_fenced_source(
        self,
        paragraph: Any,
        record: FencedSourceIdentityV3 | dict[str, Any],
        *,
        logical_body: str,
    ) -> None:
        """Bind exact fence framing without admitting any v3 number/reference fact."""

        if self._finalized:
            raise ResolvedNumberingDocxError("resolved-numbering session is finalized")
        try:
            if isinstance(record, FencedSourceIdentityV3):
                identity = fenced_source_identity_from_mapping_v3(fenced_source_mapping_v3(record))
                if identity != record:
                    raise DocxSemanticsV3Error("fenced-source identity is not canonically derived")
            else:
                identity = fenced_source_identity_from_mapping_v3(record)
            if identity.source_sha256 != self._port.source_sha256:
                raise DocxSemanticsV3Error("fenced-source record belongs to another authored source")
            if identity.source_end > len(self._port.document.authored_markdown):
                raise DocxSemanticsV3Error("fenced-source range exceeds the authored source")
            if hashlib.sha256(logical_body.encode()).hexdigest() != identity.body_sha256:
                raise DocxSemanticsV3Error("fenced-source body differs from its closed hash")
            authored = reconstruct_fenced_source_v3(identity, logical_body)
            if self._port.document.authored_markdown[identity.source_start : identity.source_end] != authored:
                raise DocxSemanticsV3Error("fenced-source reconstruction differs from authored source")
            if any(item.identity.tag == identity.tag for item in self._fenced_sources):
                raise DocxSemanticsV3Error("fenced-source tag collision")
            self._validate_direct_body_paragraph(paragraph, purpose="fenced source")
            write_canonical_fenced_body_v3(paragraph._p, logical_body)
            carrier = wrap_fenced_paragraph_payload(paragraph._p, identity.tag)
            if carrier.getparent() is not paragraph._p:
                raise DocxSemanticsV3Error("fenced-source carrier is detached from its paragraph")
            prove_fenced_source_paragraph(paragraph._p, identity)
        except DocxSemanticsV3Error as exc:
            raise ResolvedNumberingDocxError("resolved fenced-source carrier is invalid") from exc
        self._fenced_sources.append(FencedSourceBindingV3(identity, paragraph._p))

    def bind_heading(self, paragraph: Any, *, source_start: int, source_end: int) -> None:
        document_target, plan_target = self._claim_target(source_start, source_end, "heading")
        if document_target.heading_level not in self._heading_style_ids:
            raise ResolvedNumberingDocxError("Heading target has no resolved managed style")
        expected_style_id = self._heading_style_ids[document_target.heading_level]
        if getattr(getattr(paragraph, "style", None), "style_id", None) != expected_style_id:
            raise ResolvedNumberingDocxError("Heading paragraph style differs from its resolved managed style")
        rendered_text = visible_text(paragraph._p)
        key = (source_start, source_end, "heading")
        has_nested_inline = self._target_has_nested_inline(source_start, source_end)
        if not has_nested_inline and not rendered_text.startswith(document_target.authored_text):
            raise ResolvedNumberingDocxError("Heading rendered text does not preserve its authored title prefix")
        if has_nested_inline:
            fragments = self._inline_payload_fragments(
                paragraph,
                document_target,
                purpose="Heading",
                allow_rendered_suffix=True,
            )
            snapshot = tuple(element for fragment in fragments for element in fragment.elements)
            self._source_projected_heading_keys.add(key)
        else:
            from docx.oxml.ns import qn

            snapshot = tuple(deepcopy(item) for item in paragraph._p if item.tag != qn("w:pPr"))
        self._heading_rendered_texts[key] = rendered_text
        self._heading_payload_snapshots[key] = snapshot
        if not plan_target.enabled:
            self._require_style_without_numbering(expected_style_id)
        if document_target.target_id is not None:
            identity = self._target_identity(document_target, plan_target)
            wrap_paragraph_content_with_bookmark(
                paragraph,
                identity.bookmark_name,
                self._reserve_bookmark(identity.bookmark_name),
            )
            self._target_bindings.append(TargetBindingV3(identity, (paragraph._p,)))
        apply_heading_numbering(
            paragraph,
            plan_target,
            self._projection,
            heading_style_ids=self._heading_style_ids,
        )

    def bind_caption(
        self,
        caption: Any,
        object_elements: tuple[Any, ...],
        *,
        source_start: int,
        source_end: int,
        kind: str,
    ) -> None:
        if self._target_has_nested_inline(source_start, source_end):
            raise ResolvedNumberingDocxError("caption contains resolved inline carriers; use bind_rendered_caption")
        self._bind_caption(
            caption,
            object_elements,
            source_start=source_start,
            source_end=source_end,
            kind=kind,
            payload_fragments=None,
        )

    def bind_rendered_caption(
        self,
        caption: Any,
        object_elements: tuple[Any, ...],
        *,
        source_start: int,
        source_end: int,
        kind: str,
    ) -> None:
        """Prepend numbering while preserving a range-bound rich payload.

        Callers render plain runs and invoke ``render_reference`` /
        ``render_citation`` in source order in this paragraph before binding.
        The captured fragment projection is authenticated again after reopen.
        """

        fragments = self._caption_payload_fragments(caption, source_start, source_end, kind)
        self._bind_caption(
            caption,
            object_elements,
            source_start=source_start,
            source_end=source_end,
            kind=kind,
            payload_fragments=fragments,
        )

    def _bind_caption(
        self,
        caption: Any,
        object_elements: tuple[Any, ...],
        *,
        source_start: int,
        source_end: int,
        kind: str,
        payload_fragments: tuple[ResolvedInlineFragment, ...] | None,
    ) -> None:
        if kind not in {"figure", "table", "equation", "code_block"}:
            raise ResolvedNumberingDocxError("caption kind is outside the closed set")
        document_target, plan_target = self._claim_target(source_start, source_end, kind)
        logical_objects = self._validate_caption_object_elements(kind, object_elements)
        prove_caption_paragraph_style(
            caption._p,
            kind,
            self._caption_style_bindings,
        )
        identity: TargetIdentityV3 | None = None
        bookmark_id: str | None = None
        if document_target.target_id is not None:
            identity = self._target_identity(document_target, plan_target)
            bookmark_id = self._reserve_bookmark(identity.bookmark_name)
        materialize_caption_number(
            caption,
            plan_target,
            authored_content=document_target.authored_text,
            heading_style_names=self._heading_style_names,
            bookmark_name=identity.bookmark_name if identity is not None else None,
            bookmark_id=bookmark_id,
            preserve_payload=payload_fragments is not None,
        )
        key = (source_start, source_end, kind)
        self._caption_target_keys.append(key)
        self._caption_plan_bindings.append(
            _CaptionPlanBinding(
                document_target,
                plan_target,
                len(object_elements),
                tuple(deepcopy(item) for item in logical_objects),
                payload_fragments,
            )
        )
        if identity is not None:
            physical = (*object_elements, caption._p) if kind == "figure" else (caption._p, *object_elements)
            self._target_bindings.append(TargetBindingV3(identity, physical))
        elif not plan_target.enabled:
            occurrence = derive_numbering_occurrence(
                source_sha256=self._port.source_sha256,
                source_start=source_start,
                source_end=source_end,
                kind=kind,  # type: ignore[arg-type]
                plan_sha256=self._port.plan_sha256,
            )
            self._occurrence_bindings.append(_DisabledOccurrenceBinding(occurrence, caption._p, object_elements))

    def render_reference(self, paragraph: Any, *, source_start: int, source_end: int) -> None:
        if self._finalized:
            raise ResolvedNumberingDocxError("resolved-numbering session is finalized")
        key = (source_start, source_end)
        reference = self._references.get(key)
        if reference is None or key in self._bound_references:
            raise ResolvedNumberingDocxError("reference range is missing or already bound")
        self._validate_direct_body_paragraph(paragraph, purpose="reference")
        target_key = reference.target_occurrence_key
        target = self._document_targets[target_key]
        plan_target = self._plan_targets[target_key]
        if not plan_target.enabled or plan_target.derived_number != reference.cached_number:
            raise ResolvedNumberingDocxError("reference does not name one enabled plan target")
        if target.target_id is not None:
            identity = derive_target_identity_v3(target.kind, target.target_id)
            occurrence = derive_reference_occurrence_identity_v3(
                source_sha256=self._port.source_sha256,
                source_start=source_start,
                source_end=source_end,
                authored_token=reference.authored_token,
                resolved_bookmark_name=identity.bookmark_name,
                cached_number=reference.cached_number,
            )
            paragraph._p.append(
                inline_reference_sdt(
                    occurrence.tag,
                    bookmark_name=identity.bookmark_name,
                    cached_number=reference.cached_number,
                    heading_number_only=target.kind == "heading",
                    alias=reference.alias,
                )
            )
            self._reference_occurrences.append(occurrence)
            self._stable_reference_target_ids.append(target.target_id)
        else:
            identity = derive_soft_reference_identity_v3(
                source_sha256=self._port.source_sha256,
                source_start=source_start,
                source_end=source_end,
                authored_token=reference.authored_token,
                cached_number=reference.cached_number,
            )
            paragraph._p.append(
                inline_sdt(
                    identity.tag,
                    soft_reference_visible_text(identity.authored_token, identity.cached_number),
                )
            )
            self._soft_references.append(identity)
        self._bound_references.add(key)
        self._reference_paragraphs[key] = paragraph._p

    def render_citation(self, paragraph: Any, *, source_start: int, source_end: int) -> None:
        """Render one range-bound resolved Citation without consulting a database."""

        if self._finalized:
            raise ResolvedNumberingDocxError("resolved-numbering session is finalized")
        key = (source_start, source_end)
        if key not in self._citations or key in self._bound_citations:
            raise ResolvedNumberingDocxError("citation range is missing or already bound")
        identity = self._citation_occurrences.get(key)
        if identity is None:
            raise ResolvedNumberingDocxError("citation occurrence is absent from the closed projection")
        self._validate_direct_body_paragraph(paragraph, purpose="citation")
        bookmark_id = self._reserve_bookmark(identity.bookmark_name)
        paragraph._p.append(citation_occurrence_sdt(identity, bookmark_id=bookmark_id))
        self._bound_citations.add(key)
        self._citation_paragraphs[key] = paragraph._p

    def finalize_document(self) -> None:
        if self._finalized:
            return
        expected_targets = set(self._document_targets)
        if self._bound_targets != expected_targets:
            raise ResolvedNumberingDocxError("not every resolved target was bound exactly once")
        if self._bound_references != set(self._references):
            raise ResolvedNumberingDocxError("not every resolved reference was bound exactly once")
        if self._bound_citations != set(self._citations):
            raise ResolvedNumberingDocxError("not every resolved Citation was bound exactly once")
        groups = [resolved_anchor_group_v4(item) for item in self._ordinary_anchors]
        for binding in self._target_bindings:
            target = self._document_target_for_identity(binding.identity)
            groups.append(
                resolved_semantic_group_v4(
                    role="target",
                    tag=binding.identity.tag,
                    elements=binding.elements,
                    source_start=target.source_start,
                    source_end=target.source_end,
                    source_kind=target.kind,
                    payload=binding,
                )
            )
        for binding in self._occurrence_bindings:
            physical = (
                (*binding.object_elements, binding.caption_element)
                if binding.identity.kind == "figure"
                else (binding.caption_element, *binding.object_elements)
            )
            groups.append(
                resolved_semantic_group_v4(
                    role="occurrence",
                    tag=binding.identity.tag,
                    elements=physical,
                    source_start=binding.identity.source_start,
                    source_end=binding.identity.source_end,
                    source_kind=binding.identity.kind,
                    payload=binding,
                )
            )
        try:
            body = self._document.element.body
            validate_fenced_anchor_ranges_v4(
                body,
                tuple(self._ordinary_anchors),
                tuple(self._fenced_sources),
            )
            order = order_resolved_block_carriers_v4(
                body,
                tuple(groups),
                tuple(self._ordinary_anchors),
            )
            for group in order.ordered:
                if group.role == "occurrence":
                    binding = group.payload
                    wrap_numbering_occurrence(
                        binding.caption_element,
                        binding.object_elements,
                        binding.identity,
                    )
                    continue
                owners = tuple(dict.fromkeys(self._direct_body_owner(item) for item in group.elements))
                wrap_direct_body_group(owners, group.tag)
            physical_tags = physical_resolved_block_tags_v4(body)
            if set(physical_tags) != {item.tag for item in groups}:
                raise DocxSemanticsV3Error("resolved block-carrier wrapper inventory changed during finalize")
        except DocxSemanticsV3Error as exc:
            raise ResolvedNumberingDocxError("resolved source-carrier ownership is invalid") from exc
        self._anchor_topology_edges = order.anchor_topology_edges
        self._block_parent_by_tag = dict(order.direct_parent_by_tag)
        self._block_physical_tags = physical_tags
        self._finalized = True

    def write_package(
        self,
        path: str | Path,
        *,
        pre_recovery_package_transform: Callable[[Path], None] | None = None,
    ) -> None:
        """Write numbering-owned parts and bind the final physical package.

        ``pre_recovery_package_transform`` is the closed extension point for
        request-owned ZIP parts, such as footnotes and endnotes, that must be
        present before the recovery map authenticates the package projection.
        """

        self.finalize_document()
        package_path = Path(path)
        if not package_path.parent.is_dir():
            raise ResolvedNumberingDocxError("DOCX destination directory does not exist")
        parts: list[tuple[str, bytes]] = []
        if self._target_bindings or self._ordinary_anchors:
            parts.append(
                (
                    TARGET_MAP_NAMESPACE,
                    semantic_map_xml(
                        [item.identity for item in self._target_bindings],
                        [item.identity for item in self._ordinary_anchors],
                    ),
                )
            )
        if self._anchor_topology_edges:
            parts.append(
                (
                    ANCHOR_TOPOLOGY_MAP_NAMESPACE,
                    anchor_topology_map_xml(list(self._anchor_topology_edges)),
                )
            )
        if self._fenced_sources:
            records = sorted(
                (item.identity for item in self._fenced_sources),
                key=lambda item: (item.source_start, item.source_end, item.tag),
            )
            parts.append((FENCED_SOURCE_MAP_NAMESPACE, fenced_source_map_xml(records)))
        if self._soft_references:
            parts.append((SOFT_REFERENCE_MAP_NAMESPACE, soft_reference_map_xml(self._soft_references)))
        if self._reference_occurrences:
            parts.append(
                (
                    REFERENCE_OCCURRENCE_MAP_NAMESPACE,
                    reference_occurrence_map_xml(self._reference_occurrences),
                )
            )
        if self._caption_target_keys:
            parts.append(
                (
                    CAPTION_STYLE_BINDING_MAP_NAMESPACE,
                    caption_style_binding_map_xml(self._caption_style_bindings),
                )
            )
        if self._occurrence_bindings:
            parts.append(
                (
                    NUMBERING_OCCURRENCE_MAP_NAMESPACE,
                    numbering_occurrence_map_xml([item.identity for item in self._occurrence_bindings]),
                )
            )
        if self._citation_projection is not None:
            parts.extend(
                (
                    (
                        CITATION_ITEM_MAP_NAMESPACE,
                        citation_item_map_xml(self._citation_projection.item_map),
                    ),
                    (
                        CITATION_OCCURRENCE_MAP_NAMESPACE,
                        citation_occurrence_map_xml(self._citation_projection.occurrence_map),
                    ),
                )
            )
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{package_path.name}.", suffix=".tmp", dir=package_path.parent
        )
        os.close(descriptor)
        temporary = Path(temporary_name)
        try:
            self._document.save(temporary)
            write_heading_numbering_projection(temporary, self._projection)
            if parts:
                inject_custom_xml_parts(temporary, parts)
            if pre_recovery_package_transform is not None:
                pre_recovery_package_transform(temporary)
                if not temporary.is_file() or temporary.is_symlink():
                    raise ResolvedNumberingDocxError("pre-recovery package transform replaced the package unsafely")
            if self._recovery_input is not None:
                self._write_recovery_map(temporary)
            self.prove_package(temporary)
            os.replace(temporary, package_path)
        finally:
            temporary.unlink(missing_ok=True)

    def _write_recovery_map(self, temporary: Path) -> None:
        recovery_input = self._recovery_input
        if recovery_input is None:
            raise ResolvedNumberingDocxError("resolved-v4 recovery input is unavailable")
        physical_sha256 = compute_physical_projection(temporary, exclude_item_numbers=set())
        value = build_recovery_map(
            self._port,
            input_bytes=recovery_input,
            physical_sha256=physical_sha256,
        )
        inject_recovery_map(temporary, value)
        self._recovery_map = value

    def _target_has_nested_inline(self, source_start: int, source_end: int) -> bool:
        found = False
        for item in (*self._port.document.references, *self._port.document.citations):
            overlaps = item.source_start < source_end and item.source_end > source_start
            if not overlaps:
                continue
            if item.source_start < source_start or item.source_end > source_end:
                raise ResolvedNumberingDocxError("target and inline source ranges partially overlap")
            found = True
        return found

    def _caption_payload_fragments(
        self,
        caption: Any,
        source_start: int,
        source_end: int,
        kind: str,
    ) -> tuple[ResolvedInlineFragment, ...]:
        target = self._document_targets.get((source_start, source_end, kind))
        if target is None:
            raise ResolvedNumberingDocxError("rich caption target is absent from the typed snapshot")
        return self._inline_payload_fragments(
            caption,
            target,
            purpose="rich caption",
            allow_rendered_suffix=False,
        )

    def _inline_payload_fragments(
        self,
        paragraph: Any,
        target: ResolvedDocumentTarget,
        *,
        purpose: str,
        allow_rendered_suffix: bool,
    ) -> tuple[ResolvedInlineFragment, ...]:
        self._validate_direct_body_paragraph(paragraph, purpose=purpose)
        reference_tags = {
            (item.source_start, item.source_end): item.tag
            for item in (*self._soft_references, *self._reference_occurrences)
        }
        records: list[ResolvedInlineCarrier] = []
        for reference in self._port.document.references:
            overlaps = reference.source_start < target.source_end and reference.source_end > target.source_start
            if not overlaps:
                continue
            key = (reference.source_start, reference.source_end)
            if reference.source_start < target.source_start or reference.source_end > target.source_end:
                raise ResolvedNumberingDocxError(f"{purpose} and reference source ranges partially overlap")
            tag = reference_tags.get(key)
            if tag is None or self._reference_paragraphs.get(key) is not paragraph._p:
                raise ResolvedNumberingDocxError(f"{purpose} reference is unbound or belongs to another paragraph")
            records.append(ResolvedInlineCarrier(*key, tag, reference.authored_token))
        for citation in self._port.document.citations:
            overlaps = citation.source_start < target.source_end and citation.source_end > target.source_start
            if not overlaps:
                continue
            key = (citation.source_start, citation.source_end)
            if citation.source_start < target.source_start or citation.source_end > target.source_end:
                raise ResolvedNumberingDocxError(f"{purpose} and Citation source ranges partially overlap")
            identity = self._citation_occurrences.get(key)
            if identity is None or self._citation_paragraphs.get(key) is not paragraph._p:
                raise ResolvedNumberingDocxError(f"{purpose} Citation is unbound or belongs to another paragraph")
            records.append(ResolvedInlineCarrier(*key, identity.tag, citation.authored_token))
        try:
            return capture_resolved_inline_projection(
                paragraph._p,
                authored_source=self._port.document.authored_markdown,
                target_start=target.source_start,
                target_end=target.source_end,
                authored_title=target.authored_text,
                carriers=tuple(records),
                allow_rendered_suffix=allow_rendered_suffix,
            )
        except ResolvedInlineProjectionError as exc:
            raise ResolvedNumberingDocxError(f"{purpose} source projection is invalid: {exc}") from exc

    def _claim_target(
        self,
        source_start: int,
        source_end: int,
        kind: str,
    ) -> tuple[ResolvedDocumentTarget, NumberingTarget]:
        if self._finalized:
            raise ResolvedNumberingDocxError("resolved-numbering session is finalized")
        key = (source_start, source_end, kind)
        if key in self._bound_targets:
            raise ResolvedNumberingDocxError("target occurrence is already bound")
        document_target = self._document_targets.get(key)  # type: ignore[arg-type]
        plan_target = self._plan_targets.get(key)  # type: ignore[arg-type]
        if document_target is None or plan_target is None:
            raise ResolvedNumberingDocxError("target occurrence is absent from the typed snapshot")
        self._bound_targets.add(key)
        return document_target, plan_target

    def _target_identity(
        self,
        document_target: ResolvedDocumentTarget,
        plan_target: NumberingTarget,
    ) -> TargetIdentityV3:
        source_id = document_target.target_id
        if source_id is None or source_id != plan_target.target_id:
            raise ResolvedNumberingDocxError("target identity contradicts its plan")
        self._claim_source_id(source_id)
        return derive_target_identity_v3(document_target.kind, source_id)

    def _document_target_for_identity(self, identity: TargetIdentityV3) -> ResolvedDocumentTarget:
        matches = [
            item
            for item in self._document_targets.values()
            if item.kind == identity.kind and item.target_id == identity.source_id
        ]
        if len(matches) != 1:
            raise ResolvedNumberingDocxError("resolved target identity has no unique typed occurrence")
        return matches[0]

    def _claim_source_id(self, source_id: str) -> None:
        if source_id in self._claimed_ids:
            raise ResolvedNumberingDocxError("resolved target/anchor source ID is duplicated")
        self._claimed_ids.add(source_id)

    def _reserve_bookmark(self, bookmark_name: str) -> str:
        inventory = build_docx_bookmark_inventory(self._document)
        if bookmark_name.casefold() in inventory.used_name_keys:
            raise ResolvedNumberingDocxError("resolved target bookmark conflicts with the template")
        used_ids = set(inventory.used_id_keys)
        for candidate in range(BOOKMARK_ID_MAX + 1):
            raw = str(candidate)
            if bookmark_id_key(raw) not in used_ids:
                return raw
        raise ResolvedNumberingDocxError("no portable DOCX bookmark IDs remain")

    def _direct_body_owner(self, element: Any) -> Any:
        body = self._document.element.body
        owner = element
        while owner.getparent() is not body:
            owner = owner.getparent()
            if owner is None:
                raise ResolvedNumberingDocxError("resolved logical block is detached from body")
        return owner

    def _validate_direct_body_paragraph(self, paragraph: Any, *, purpose: str) -> None:
        if paragraph.part is not self._document.part:
            raise ResolvedNumberingDocxError(f"{purpose} paragraph belongs to another document part")
        if paragraph._p.getparent() is not self._document.element.body:
            raise ResolvedNumberingDocxError(f"{purpose} paragraph must be a direct main-body paragraph")


__all__ = ["ResolvedNumberingDocxError", "ResolvedNumberingDocxSession"]
