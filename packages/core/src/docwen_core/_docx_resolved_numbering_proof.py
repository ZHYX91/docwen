"""Reopen proof for a request-owned resolved-numbering DOCX session."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from zipfile import ZipFile

from docwen_core._docx_caption_carrier import captionable_logical_elements
from docwen_core._docx_numbering_package import _elements_equal
from docwen_core._docx_resolved_source_carriers import (
    prove_resolved_anchor_topology_v4,
    prove_resolved_block_hierarchy_v4,
    prove_resolved_ordinary_anchor_group_v4,
)
from docwen_core._docx_semantics_v3_fenced import (
    FENCED_SOURCE_MAP_NAMESPACE,
    bind_fenced_source_document_v3,
)
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
    DocxSemanticsV3Error,
    ReferenceOccurrenceIdentityV3,
    derive_target_identity_v3,
)
from docwen_core._docx_semantics_v3_ooxml import (
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
    parse_caption_style_binding_map,
    prove_caption_paragraph_style,
    prove_caption_style_registry,
)
from docwen_core._docx_semantics_v3_topology import (
    parse_anchor_topology_map,
    prove_alias_run,
    prove_block_sdt_envelope,
    prove_exact_field_run,
    prove_inline_occurrence_envelope,
    prove_soft_reference_envelope,
)
from docwen_core.docx_bookmarks import (
    build_docx_bookmark_inventory,
    prove_bookmark_name,
)
from docwen_core.docx_citation_ooxml import (
    CITATION_ITEM_MAP_NAMESPACE,
    CITATION_OCCURRENCE_MAP_NAMESPACE,
    parse_citation_item_map,
    parse_citation_occurrence_map,
    prove_citation_projection,
    validate_citation_authorities,
)
from docwen_core.docx_numbering_import import HeadingNumberingProofIndex
from docwen_core.docx_numbering_occurrence import (
    NUMBERING_OCCURRENCE_MAP_NAMESPACE,
    parse_numbering_occurrence_map,
    prove_numbering_occurrence_sdt,
)
from docwen_core.docx_numbering_ooxml import (
    ResolvedNumberingOoxmlError,
    prove_caption_number,
    prove_heading_numbering_projection,
)


class ResolvedNumberingDocxError(ResolvedNumberingOoxmlError):
    """The validated port and rendered document no longer agree."""


class ResolvedNumberingProofMixin:
    """Authenticate every session-owned physical carrier after package reopen."""

    _caption_plan_bindings: list[Any]
    _caption_style_bindings: tuple[Any, ...]
    _caption_target_keys: list[Any]
    _citation_projection: Any | None
    _document: Any
    _document_targets: dict[Any, Any]
    _finalized: bool
    _heading_rendered_texts: dict[Any, str]
    _heading_payload_snapshots: dict[Any, tuple[Any, ...]]
    _heading_style_ids: dict[int, str]
    _heading_style_names: dict[str, str]
    _occurrence_bindings: list[Any]
    _ordinary_anchors: list[Any]
    _anchor_topology_edges: tuple[Any, ...]
    _fenced_sources: list[Any]
    _block_parent_by_tag: dict[str, str | None]
    _block_physical_tags: tuple[str, ...]
    _plan_targets: dict[Any, Any]
    _projection: Any
    _reference_occurrences: list[Any]
    _references: dict[Any, Any]
    _soft_references: list[Any]
    _source_projected_heading_keys: set[Any]
    _target_bindings: list[Any]

    def prove_package(self, path: str | Path) -> None:
        """Reopen the saved package and authenticate the new occurrence layer."""

        if not self._finalized:
            raise ResolvedNumberingDocxError("session must be finalized before package proof")
        from docx import Document

        package_path = Path(path)
        with ZipFile(package_path) as package:
            owned = read_owned_map_parts(package)
            for namespace, (item_number, _root) in owned.items():
                verify_custom_xml_support(package, item_number, namespace)
            self._prove_target_map(owned)
            self._prove_anchor_topology_map(owned)
            self._prove_fenced_source_map(owned)
            self._prove_soft_reference_map(owned)
            self._prove_reference_occurrence_map(owned)
            self._prove_caption_style_map(package, owned)
            self._prove_numbering_occurrence_map(owned)
            citation_namespaces = {
                CITATION_ITEM_MAP_NAMESPACE,
                CITATION_OCCURRENCE_MAP_NAMESPACE,
            }
            present_citation_namespaces = citation_namespaces.intersection(owned)
            if self._citation_projection is None:
                if present_citation_namespaces:
                    raise ResolvedNumberingDocxError("unexpected resolved-citation map is present")
            elif present_citation_namespaces != citation_namespaces:
                raise ResolvedNumberingDocxError("resolved Citation does not have exactly both authority maps")
            else:
                _item_number, item_root = owned[CITATION_ITEM_MAP_NAMESPACE]
                _occurrence_number, occurrence_root = owned[CITATION_OCCURRENCE_MAP_NAMESPACE]
                item_map = parse_citation_item_map(item_root)
                occurrence_map = parse_citation_occurrence_map(occurrence_root)
                validate_citation_authorities(item_map, occurrence_map)
                if (
                    item_map != self._citation_projection.item_map
                    or occurrence_map != self._citation_projection.occurrence_map
                ):
                    raise ResolvedNumberingDocxError("resolved-citation maps differ after reopen")
        reopened = Document(str(package_path))
        prove_heading_numbering_projection(package_path, self._projection)
        bookmark_inventory = build_docx_bookmark_inventory(reopened)
        self._prove_target_sdts(reopened, bookmark_inventory)
        self._prove_heading_paragraphs(package_path, reopened, bookmark_inventory)
        caption_paragraphs = self._prove_caption_paragraphs(reopened, bookmark_inventory)
        self._prove_inline_reference_sdts(reopened)
        self._prove_numbering_occurrence_sdts(reopened, caption_paragraphs)
        if self._citation_projection is not None:
            prove_citation_projection(reopened, self._citation_projection)
        self._prove_source_carriers(reopened, set(caption_paragraphs.values()))
        self._prove_merged_inline_order(reopened)

    def _prove_target_map(self, owned: dict[str, tuple[int, Any]]) -> None:
        if not self._target_bindings and not self._ordinary_anchors:
            if TARGET_MAP_NAMESPACE in owned:
                raise ResolvedNumberingDocxError("unexpected target map is present")
            return
        if TARGET_MAP_NAMESPACE not in owned:
            raise ResolvedNumberingDocxError("target map is missing")
        _number, root = owned[TARGET_MAP_NAMESPACE]
        targets, anchors = parse_semantic_map(root)
        expected = sorted(
            (item.identity for item in self._target_bindings),
            key=lambda item: (item.kind, item.source_id),
        )
        expected_anchors = sorted(
            (item.identity for item in self._ordinary_anchors),
            key=lambda item: (item.block_kind, item.source_id),
        )
        if targets != expected or anchors != expected_anchors:
            raise ResolvedNumberingDocxError("target map differs after reopen")

    def _prove_anchor_topology_map(self, owned: dict[str, tuple[int, Any]]) -> None:
        if not self._anchor_topology_edges:
            if ANCHOR_TOPOLOGY_MAP_NAMESPACE in owned:
                raise ResolvedNumberingDocxError("unexpected ordinary-anchor topology map is present")
            return
        if ANCHOR_TOPOLOGY_MAP_NAMESPACE not in owned:
            raise ResolvedNumberingDocxError("ordinary-anchor topology map is missing")
        _number, root = owned[ANCHOR_TOPOLOGY_MAP_NAMESPACE]
        if tuple(parse_anchor_topology_map(root)) != self._anchor_topology_edges:
            raise ResolvedNumberingDocxError("ordinary-anchor topology map differs after reopen")

    def _prove_fenced_source_map(self, owned: dict[str, tuple[int, Any]]) -> None:
        if not self._fenced_sources:
            if FENCED_SOURCE_MAP_NAMESPACE in owned:
                raise ResolvedNumberingDocxError("unexpected fenced-source map is present")
            return
        if FENCED_SOURCE_MAP_NAMESPACE not in owned:
            raise ResolvedNumberingDocxError("fenced-source map is missing")
        _number, root = owned[FENCED_SOURCE_MAP_NAMESPACE]
        expected = sorted(
            (item.identity for item in self._fenced_sources),
            key=lambda item: (item.source_start, item.source_end, item.tag),
        )
        if parse_fenced_source_map(root) != expected:
            raise ResolvedNumberingDocxError("fenced-source map differs after reopen")

    def _prove_soft_reference_map(self, owned: dict[str, tuple[int, Any]]) -> None:
        if not self._soft_references:
            if SOFT_REFERENCE_MAP_NAMESPACE in owned:
                raise ResolvedNumberingDocxError("unexpected soft-reference map is present")
            return
        if SOFT_REFERENCE_MAP_NAMESPACE not in owned:
            raise ResolvedNumberingDocxError("soft-reference map is missing")
        _number, root = owned[SOFT_REFERENCE_MAP_NAMESPACE]
        expected = sorted(
            self._soft_references,
            key=lambda item: (item.source_start, item.source_end, item.tag),
        )
        if parse_soft_reference_map(root) != expected:
            raise ResolvedNumberingDocxError("soft-reference map differs after reopen")

    def _prove_reference_occurrence_map(self, owned: dict[str, tuple[int, Any]]) -> None:
        if not self._reference_occurrences:
            if REFERENCE_OCCURRENCE_MAP_NAMESPACE in owned:
                raise ResolvedNumberingDocxError("unexpected reference-occurrence map is present")
            return
        if REFERENCE_OCCURRENCE_MAP_NAMESPACE not in owned:
            raise ResolvedNumberingDocxError("reference-occurrence map is missing")
        _number, root = owned[REFERENCE_OCCURRENCE_MAP_NAMESPACE]
        expected = sorted(
            self._reference_occurrences,
            key=lambda item: (item.source_start, item.source_end, item.tag),
        )
        if parse_reference_occurrence_map(root) != expected:
            raise ResolvedNumberingDocxError("reference-occurrence map differs after reopen")

    def _prove_caption_style_map(
        self,
        package: ZipFile,
        owned: dict[str, tuple[int, Any]],
    ) -> None:
        if not self._caption_target_keys:
            if CAPTION_STYLE_BINDING_MAP_NAMESPACE in owned:
                raise ResolvedNumberingDocxError("unexpected caption-style map is present")
            return
        if CAPTION_STYLE_BINDING_MAP_NAMESPACE not in owned:
            raise ResolvedNumberingDocxError("caption-style map is missing")
        _number, root = owned[CAPTION_STYLE_BINDING_MAP_NAMESPACE]
        if parse_caption_style_binding_map(root) != self._caption_style_bindings:
            raise ResolvedNumberingDocxError("caption-style map differs after reopen")
        prove_caption_style_registry(package, self._caption_style_bindings)

    def _prove_numbering_occurrence_map(self, owned: dict[str, tuple[int, Any]]) -> None:
        if not self._occurrence_bindings:
            if NUMBERING_OCCURRENCE_MAP_NAMESPACE in owned:
                raise ResolvedNumberingDocxError("unexpected numbering-occurrence map is present")
            return
        if NUMBERING_OCCURRENCE_MAP_NAMESPACE not in owned:
            raise ResolvedNumberingDocxError("numbering-occurrence map is missing")
        _number, root = owned[NUMBERING_OCCURRENCE_MAP_NAMESPACE]
        expected = sorted(
            (item.identity for item in self._occurrence_bindings),
            key=lambda item: (item.source_start, item.source_end, item.tag),
        )
        if parse_numbering_occurrence_map(root) != expected:
            raise ResolvedNumberingDocxError("numbering-occurrence map differs after reopen")

    def _prove_target_sdts(self, document: Any, bookmark_inventory: Any) -> None:
        from docx.oxml.ns import qn

        body = document.element.body
        expected_by_tag = {item.identity.tag: item.identity for item in self._target_bindings}
        found: dict[str, list[Any]] = {tag: [] for tag in expected_by_tag}
        for sdt in body.iter(qn("w:sdt")):
            tag = sdt_tag(sdt) or ""
            if not tag.startswith("docwen-target-v1:"):
                continue
            if tag not in expected_by_tag:
                raise ResolvedNumberingDocxError("unmapped target SDT is present")
            found[tag].append(sdt)
        for tag, identity in expected_by_tag.items():
            candidates = found[tag]
            if len(candidates) != 1:
                raise ResolvedNumberingDocxError("target SDT is missing or duplicated")
            sdt = candidates[0]
            parent = sdt.getparent()
            if parent is not body and (
                parent is None
                or parent.tag != qn("w:sdtContent")
                or not (sdt_tag(parent.getparent()) or "").startswith(ANCHOR_TAG_PREFIX)
            ):
                raise ResolvedNumberingDocxError("target SDT has an unsupported block owner")
            prove_block_sdt_envelope(sdt, tag)
            proof = prove_bookmark_name(bookmark_inventory, identity.bookmark_name, scope_element=sdt)
            if not proof.valid:
                raise ResolvedNumberingDocxError("target bookmark is not uniquely bound inside its SDT")

    def _prove_heading_paragraphs(self, path: Path, document: Any, bookmark_inventory: Any) -> None:
        from docx.oxml.ns import qn
        from docx.text.paragraph import Paragraph

        bindings = sorted(
            (
                (target, self._plan_targets[target.occurrence_key])
                for target in self._document_targets.values()
                if target.kind == "heading"
            ),
            key=lambda item: (item[0].source_start, item[0].source_end),
        )
        heading_style_ids = set(self._heading_style_ids.values())
        paragraphs = [
            item
            for item in document.element.body.iter(qn("w:p"))
            if self._direct_paragraph_style_id(item) in heading_style_ids
        ]
        if len(paragraphs) != len(bindings):
            raise ResolvedNumberingDocxError("Heading paragraph inventory differs from the typed snapshot")
        proof_index = HeadingNumberingProofIndex.load(path)
        for paragraph_element, (document_target, plan_target) in zip(paragraphs, bindings, strict=True):
            level = document_target.heading_level
            if level is None or self._direct_paragraph_style_id(paragraph_element) != self._heading_style_ids[level]:
                raise ResolvedNumberingDocxError("Heading paragraph style differs after reopen")
            expected_text = self._heading_rendered_texts.get(document_target.occurrence_key)
            if expected_text is None or (
                document_target.occurrence_key not in self._source_projected_heading_keys
                and not expected_text.startswith(document_target.authored_text)
            ):
                raise ResolvedNumberingDocxError("Heading render snapshot is missing or invalid")
            if visible_text(paragraph_element) != expected_text:
                raise ResolvedNumberingDocxError("Heading rendered text changed during materialization")
            payload = tuple(item for item in paragraph_element if item.tag != qn("w:pPr"))
            if document_target.target_id is not None:
                payload = payload[1:-1]
            expected_payload = self._heading_payload_snapshots.get(document_target.occurrence_key)
            if (
                expected_payload is None
                or len(payload) != len(expected_payload)
                or any(
                    not _elements_equal(actual, expected)
                    for actual, expected in zip(payload, expected_payload, strict=True)
                )
            ):
                raise ResolvedNumberingDocxError("Heading authored OOXML payload changed after binding")
            paragraph = Paragraph(paragraph_element, document)
            proof = proof_index.prove_paragraph(paragraph)
            direct_num_pr = paragraph_element.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
            if not plan_target.enabled:
                if direct_num_pr is not None or proof is not None:
                    raise ResolvedNumberingDocxError("disabled Heading retained effective numbering")
            else:
                materialization = plan_target.materialization
                if materialization is None or getattr(materialization, "type", None) != "heading_list" or proof is None:
                    raise ResolvedNumberingDocxError("enabled Heading lacks proven list semantics")
                if (
                    proof.num_id != self._projection.num_id(materialization.instance_id)
                    or proof.abstract_num_id != self._projection.abstract_id(materialization.definition_id)
                    or proof.level != materialization.level
                    or proof.style_id != self._heading_style_ids[materialization.level]
                ):
                    raise ResolvedNumberingDocxError("Heading list binding differs from its closed plan")
            if document_target.target_id is not None:
                identity = derive_target_identity_v3("heading", document_target.target_id)
                self._prove_heading_bookmark(paragraph_element, identity.bookmark_name, bookmark_inventory)
                owner = paragraph_element.getparent()
                sdt = None if owner is None else owner.getparent()
                if owner is None or owner.tag != qn("w:sdtContent") or sdt_tag(sdt) != identity.tag:
                    raise ResolvedNumberingDocxError("addressable Heading is outside its target SDT")

    def _prove_caption_paragraphs(
        self,
        document: Any,
        bookmark_inventory: Any,
    ) -> dict[tuple[int, int, str], Any]:
        from docx.oxml.ns import qn

        bindings = sorted(
            self._caption_plan_bindings,
            key=lambda item: (item.document_target.source_start, item.document_target.source_end),
        )
        style_by_kind = self._caption_style_ids_by_kind()
        caption_style_ids = set(style_by_kind.values())
        paragraphs = [
            item
            for item in document.element.body.iter(qn("w:p"))
            if self._direct_paragraph_style_id(item) in caption_style_ids
        ]
        if len(paragraphs) != len(bindings):
            raise ResolvedNumberingDocxError("caption paragraph inventory differs from the typed snapshot")
        proven: dict[tuple[int, int, str], Any] = {}
        for paragraph, binding in zip(paragraphs, bindings, strict=True):
            document_target = binding.document_target
            plan_target = binding.plan_target
            if self._direct_paragraph_style_id(paragraph) != style_by_kind[document_target.kind]:
                raise ResolvedNumberingDocxError("caption kind/style binding differs after reopen")
            prove_caption_paragraph_style(paragraph, document_target.kind, self._caption_style_bindings)
            bookmark_name = (
                derive_target_identity_v3(document_target.kind, document_target.target_id).bookmark_name
                if document_target.target_id is not None
                else None
            )
            rendered_payload = (
                None
                if binding.payload_fragments is None
                else tuple(element for fragment in binding.payload_fragments for element in fragment.elements)
            )
            prove_caption_number(
                paragraph,
                plan_target,
                authored_content=document_target.authored_text,
                heading_style_names=self._heading_style_names,
                bookmark_name=bookmark_name,
                bookmark_inventory=bookmark_inventory,
                rendered_payload=rendered_payload,
            )
            self._prove_caption_adjacency(paragraph, binding)
            owner = paragraph.getparent()
            sdt = None if owner is None else owner.getparent()
            if document_target.target_id is not None:
                identity = derive_target_identity_v3(document_target.kind, document_target.target_id)
                if owner is None or owner.tag != qn("w:sdtContent") or sdt_tag(sdt) != identity.tag:
                    raise ResolvedNumberingDocxError("addressable caption is outside its target SDT")
            elif plan_target.enabled:
                if owner is document.element.body:
                    pass
                elif (
                    owner is None
                    or owner.tag != qn("w:sdtContent")
                    or not (sdt_tag(owner.getparent()) or "").startswith(ANCHOR_TAG_PREFIX)
                ):
                    raise ResolvedNumberingDocxError("enabled ID-less caption has an unexpected ownership wrapper")
            proven[document_target.occurrence_key] = paragraph
        return proven

    def _prove_numbering_occurrence_sdts(
        self,
        document: Any,
        caption_paragraphs: dict[tuple[int, int, str], Any],
    ) -> None:
        from docx.oxml.ns import qn

        expected = sorted(
            self._occurrence_bindings,
            key=lambda item: (item.identity.source_start, item.identity.source_end, item.identity.tag),
        )
        body = document.element.body
        all_occurrences = [
            sdt for sdt in body.iter(qn("w:sdt")) if (sdt_tag(sdt) or "").startswith("docwen-numbering-occurrence-v1:")
        ]
        expected_tags = [item.identity.tag for item in expected]
        if [sdt_tag(item) for item in all_occurrences] != expected_tags:
            raise ResolvedNumberingDocxError(
                "numbering-occurrence physical order/cardinality differs from source authority"
            )
        style_by_kind = {
            "figure": "figure_caption",
            "table": "table_caption",
            "equation": "equation_caption",
            "code_block": "code_block_caption",
        }
        style_ids = {item.semantic_key: item.resolved_style_id for item in self._caption_style_bindings}
        binding_by_key = {item.document_target.occurrence_key: item for item in self._caption_plan_bindings}
        for sdt, occurrence in zip(all_occurrences, expected, strict=True):
            parent = sdt.getparent()
            if parent is not body and (
                parent is None
                or parent.tag != qn("w:sdtContent")
                or not (sdt_tag(parent.getparent()) or "").startswith(ANCHOR_TAG_PREFIX)
            ):
                raise ResolvedNumberingDocxError("numbering occurrence has an unsupported block owner")
            key = (
                occurrence.identity.source_start,
                occurrence.identity.source_end,
                occurrence.identity.kind,
            )
            binding = binding_by_key[key]
            inline_tags = (
                ()
                if binding.payload_fragments is None
                else tuple(
                    fragment.carrier_tag for fragment in binding.payload_fragments if fragment.carrier_tag is not None
                )
            )
            caption, _logical_object = prove_numbering_occurrence_sdt(
                sdt,
                occurrence.identity,
                caption_style_id=style_ids[style_by_kind[occurrence.identity.kind]],
                allowed_inline_tags=inline_tags,
            )
            if caption is not caption_paragraphs.get(key):
                raise ResolvedNumberingDocxError("numbering-occurrence tag is not bound to its exact caption paragraph")

    def _prove_source_carriers(self, document: Any, caption_paragraphs: set[Any]) -> None:
        """Reopen and authenticate the carrier-only source projection."""

        expected_tags = {
            *(item.identity.tag for item in self._target_bindings),
            *(item.identity.tag for item in self._ordinary_anchors),
            *(item.identity.tag for item in self._occurrence_bindings),
        }
        try:
            wrappers = prove_resolved_block_hierarchy_v4(
                document.element.body,
                expected_tags=expected_tags,
                expected_physical_tags=self._block_physical_tags,
                expected_parent_by_tag=self._block_parent_by_tag,
            )
            prove_resolved_anchor_topology_v4(
                tuple(self._ordinary_anchors),
                self._anchor_topology_edges,
                wrappers,
            )
            for binding in self._ordinary_anchors:
                prove_resolved_ordinary_anchor_group_v4(
                    wrappers[binding.identity.tag],
                    binding,
                    allowed_caption_paragraphs=caption_paragraphs,
                )
            records = sorted(
                (item.identity for item in self._fenced_sources),
                key=lambda item: (item.source_start, item.source_end, item.tag),
            )
            if records:
                bound = bind_fenced_source_document_v3(document.element.body, records)
                if len(bound) != len(records):  # pragma: no cover - binder already owns cardinality
                    raise DocxSemanticsV3Error("fenced-source reopen inventory differs from authority")
        except DocxSemanticsV3Error as exc:
            raise ResolvedNumberingDocxError("resolved source-carrier reopen proof failed") from exc

    def _prove_inline_reference_sdts(self, document: Any) -> None:
        from docx.oxml.ns import qn

        body = document.element.body
        soft_by_tag = {item.tag: item for item in self._soft_references}
        stable_by_tag = {item.tag: item for item in self._reference_occurrences}
        soft_physical: list[str] = []
        stable_physical: list[str] = []
        for sdt in body.iter(qn("w:sdt")):
            tag = sdt_tag(sdt) or ""
            if tag.startswith(SOFT_REFERENCE_TAG_PREFIX):
                record = soft_by_tag.get(tag)
                if record is None or tag in soft_physical or sdt.getparent().tag != qn("w:p"):
                    raise ResolvedNumberingDocxError("soft-reference SDT is duplicated, unmapped, or misplaced")
                prove_soft_reference_envelope(
                    sdt,
                    tag,
                    soft_reference_visible_text(record.authored_token, record.cached_number),
                )
                soft_physical.append(tag)
            elif tag.startswith(REFERENCE_OCCURRENCE_TAG_PREFIX):
                record = stable_by_tag.get(tag)
                if record is None or tag in stable_physical or sdt.getparent().tag != qn("w:p"):
                    raise ResolvedNumberingDocxError("reference-occurrence SDT is duplicated, unmapped, or misplaced")
                self._prove_stable_reference_sdt(sdt, record)
                stable_physical.append(tag)
        expected_soft = [
            item.tag
            for item in sorted(self._soft_references, key=lambda item: (item.source_start, item.source_end, item.tag))
        ]
        expected_stable = [
            item.tag
            for item in sorted(
                self._reference_occurrences,
                key=lambda item: (item.source_start, item.source_end, item.tag),
            )
        ]
        if soft_physical != expected_soft or stable_physical != expected_stable:
            raise ResolvedNumberingDocxError("reference physical order differs from source authority")

    def _prove_merged_inline_order(self, document: Any) -> None:
        from docx.oxml.ns import qn

        expected: list[tuple[int, int, str]] = [
            (item.source_start, item.source_end, item.tag)
            for item in (*self._soft_references, *self._reference_occurrences)
        ]
        if self._citation_projection is not None:
            expected.extend(
                (item.source_start, item.source_end, item.tag)
                for item in self._citation_projection.occurrence_map.occurrences
            )
        expected_tags = [item[2] for item in sorted(expected)]
        prefixes = (
            SOFT_REFERENCE_TAG_PREFIX,
            REFERENCE_OCCURRENCE_TAG_PREFIX,
            "docwen-citation-occurrence-v1:",
        )
        physical_tags = [
            value
            for item in document.element.body.iter(qn("w:tag"))
            if (value := item.get(qn("w:val"))) is not None and value.startswith(prefixes)
        ]
        if physical_tags != expected_tags:
            raise ResolvedNumberingDocxError("inline carrier physical order differs from source authority")

    def _prove_stable_reference_sdt(self, sdt: Any, record: ReferenceOccurrenceIdentityV3) -> None:
        from docx.oxml.ns import qn

        prove_inline_occurrence_envelope(sdt, record.tag)
        content = sdt.find(qn("w:sdtContent"))
        if content is None:
            raise ResolvedNumberingDocxError("reference occurrence has no content")
        children = list(content)
        reference = self._references[(record.source_start, record.source_end)]
        target = self._document_targets[reference.target_occurrence_key]
        expected_length = 6 if reference.alias is not None else 5
        if len(children) != expected_length or any(item.tag != qn("w:r") for item in children):
            raise ResolvedNumberingDocxError("reference occurrence field topology is not canonical")
        prove_exact_field_run(children[0], qn("w:fldChar"), field_type="begin", dirty="true")
        prove_exact_field_run(children[1], qn("w:instrText"), xml_space="preserve")
        prove_exact_field_run(children[2], qn("w:fldChar"), field_type="separate")
        prove_exact_field_run(children[3], qn("w:t"))
        prove_exact_field_run(children[4], qn("w:fldChar"), field_type="end")
        instruction = (
            f" REF {record.resolved_bookmark_name} \\n \\h "
            if target.kind == "heading"
            else f" REF {record.resolved_bookmark_name} \\h "
        )
        if (children[1][0].text or "") != instruction or (children[3][0].text or "") != record.cached_number:
            raise ResolvedNumberingDocxError("reference instruction or cached number differs from authority")
        if reference.alias is not None:
            prove_alias_run(children[5])
            if (children[5][0].text or "") != f" {reference.alias}":
                raise ResolvedNumberingDocxError("reference Alias differs from authored authority")

    def _prove_heading_bookmark(self, paragraph: Any, bookmark_name: str, bookmark_inventory: Any) -> None:
        from docx.oxml.ns import qn

        payload = [item for item in paragraph if item.tag != qn("w:pPr")]
        if (
            len(payload) < 2
            or payload[0].tag != qn("w:bookmarkStart")
            or payload[0].get(qn("w:name")) != bookmark_name
            or payload[-1].tag != qn("w:bookmarkEnd")
            or payload[-1].get(qn("w:id")) != payload[0].get(qn("w:id"))
        ):
            raise ResolvedNumberingDocxError("Heading bookmark does not enclose the complete authored content")
        proof = prove_bookmark_name(bookmark_inventory, bookmark_name, scope_element=paragraph)
        if not proof.valid or proof.start is None or proof.end is None:
            raise ResolvedNumberingDocxError("Heading bookmark is not globally proven")
        if proof.start.element is not payload[0] or proof.end.element is not payload[-1]:
            raise ResolvedNumberingDocxError("Heading bookmark range differs from its paragraph")

    def _prove_caption_adjacency(self, paragraph: Any, binding: Any) -> None:
        from docx.oxml.ns import qn

        container = paragraph.getparent()
        if container is None or container.tag not in {qn("w:body"), qn("w:sdtContent")}:
            raise ResolvedNumberingDocxError("caption is outside a supported block container")
        children = list(container)
        try:
            index = children.index(paragraph)
        except ValueError as exc:  # pragma: no cover - lxml parent invariant
            raise ResolvedNumberingDocxError("caption is detached from its physical container") from exc
        count = binding.object_count
        if binding.document_target.kind == "figure":
            object_elements = tuple(children[index - count : index]) if index >= count else ()
            exact_slot = index == count
        else:
            object_elements = tuple(children[index + 1 : index + 1 + count])
            exact_slot = index + 1 + count == len(children)
        if len(object_elements) != count:
            raise ResolvedNumberingDocxError("caption lost its directly adjacent logical object")
        parent_tag = sdt_tag(container.getparent()) if container.tag == qn("w:sdtContent") else None
        exact_owner = parent_tag is not None and parent_tag.startswith(
            ("docwen-target-v1:", "docwen-numbering-occurrence-v1:")
        )
        if exact_owner and (len(children) != count + 1 or not exact_slot):
            raise ResolvedNumberingDocxError("caption ownership wrapper has wrong block cardinality")
        logical_objects = self._validate_caption_object_elements(binding.document_target.kind, object_elements)
        if len(logical_objects) != len(binding.object_snapshots) or any(
            not _elements_equal(actual, expected)
            for actual, expected in zip(logical_objects, binding.object_snapshots, strict=True)
        ):
            raise ResolvedNumberingDocxError("caption logical object OOXML differs from its bind snapshot")

    def _validate_caption_object_elements(self, kind: str, elements: tuple[Any, ...]) -> tuple[Any, ...]:
        try:
            logical = captionable_logical_elements(elements)
        except DocxSemanticsV3Error as exc:
            raise ResolvedNumberingDocxError("caption logical object wrapper is invalid") from exc
        if kind not in {"figure", "table", "equation", "code_block"}:
            raise ResolvedNumberingDocxError("caption target has an unsupported semantic kind")
        return tuple(logical)

    def _caption_style_ids_by_kind(self) -> dict[str, str]:
        by_key = {item.semantic_key: item.resolved_style_id for item in self._caption_style_bindings}
        return {
            "figure": by_key["figure_caption"],
            "table": by_key["table_caption"],
            "equation": by_key["equation_caption"],
            "code_block": by_key["code_block_caption"],
        }

    def _require_style_without_numbering(self, style_id: str) -> None:
        from docx.oxml.ns import qn

        styles = self._document.styles.element.findall(qn("w:style"))
        by_id = {item.get(qn("w:styleId")): item for item in styles}
        seen: set[str] = set()
        current = style_id
        while current:
            if current in seen or current not in by_id:
                raise ResolvedNumberingDocxError("Heading style inheritance is cyclic or unresolved")
            seen.add(current)
            style = by_id[current]
            if style.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None:
                raise ResolvedNumberingDocxError("disabled Heading style has inherited numbering")
            based_on = style.findall(qn("w:basedOn"))
            if len(based_on) > 1:
                raise ResolvedNumberingDocxError("Heading style has duplicate basedOn relationships")
            current = based_on[0].get(qn("w:val"), "") if based_on else ""

    @staticmethod
    def _direct_paragraph_style_id(paragraph: Any) -> str | None:
        from docx.oxml.ns import qn

        styles = paragraph.findall(f"{qn('w:pPr')}/{qn('w:pStyle')}")
        return styles[0].get(qn("w:val")) if len(styles) == 1 else None
