"""Request-owned DOCX render session for Markdown semantics v3."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

import lxml.etree as etree

from docwen_core._docx_semantics_v3_fenced import (
    FENCED_SOURCE_MAP_NAMESPACE,
    FencedSourceBindingV3,
    FencedSourceIdentityV3,
    prove_fenced_source_paragraph,
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
    AnchorBindingV3,
    AnchorTopologyEdgeV3,
    CaptionBindingV3,
    CaptionStyleBindingV3,
    DocxSemanticsV3Error,
    ReferenceOccurrenceIdentityV3,
    SoftReferenceIdentityV3,
    TargetBindingV3,
    derive_anchor_identity_v3,
    derive_reference_occurrence_identity_v3,
    derive_soft_reference_identity_v3,
    derive_target_identity_v3,
    require_sha256,
    require_source_id,
)
from docwen_core._docx_semantics_v3_ooxml import (
    inline_reference_sdt,
    inline_sdt,
    soft_reference_visible_text,
    wrap_direct_body_group,
    wrap_paragraph_content_with_bookmark,
)
from docwen_core._docx_semantics_v3_package import (
    inject_custom_xml_parts,
    reference_occurrence_map_xml,
    semantic_map_xml,
    soft_reference_map_xml,
)
from docwen_core._docx_semantics_v3_recovery import DocxSemanticsV3Recovery
from docwen_core._docx_semantics_v3_styles import (
    caption_style_binding_map_xml,
    prove_caption_paragraph_style,
    validate_caption_style_bindings,
)
from docwen_core._docx_semantics_v3_topology import anchor_topology_map_xml, order_ordinary_anchor_bindings
from docwen_core.docx_bookmarks import (
    BOOKMARK_ID_MAX,
    bookmark_id_key,
    build_docx_bookmark_inventory,
)


class DocxSemanticsV3Session:
    """Request-owned renderer state, finalized only after template insertion."""

    def __init__(
        self,
        document: Any,
        *,
        source_sha256: str,
        caption_style_bindings: tuple[CaptionStyleBindingV3, ...] = (),
    ) -> None:
        require_sha256(source_sha256)
        self._document = document
        self._source_sha256 = source_sha256
        self._caption_style_bindings = (
            validate_caption_style_bindings(caption_style_bindings) if caption_style_bindings else ()
        )
        self._targets: list[TargetBindingV3] = []
        self._anchors: list[AnchorBindingV3] = []
        self._anchor_topology_edges: tuple[AnchorTopologyEdgeV3, ...] = ()
        self._captions: list[CaptionBindingV3] = []
        self._stable_reference_target_ids: list[str] = []
        self._reference_occurrences: list[ReferenceOccurrenceIdentityV3] = []
        self._soft_references: list[SoftReferenceIdentityV3] = []
        self._fenced_sources: list[FencedSourceBindingV3] = []
        self._owned_source_ids: set[str] = set()
        self._finalized = False

    @property
    def has_projection(self) -> bool:
        return bool(
            self._targets
            or self._anchors
            or self._captions
            or self._reference_occurrences
            or self._soft_references
            or self._fenced_sources
        )

    def bind_fenced_source(
        self,
        paragraph: Any,
        record: FencedSourceIdentityV3 | dict[str, Any],
        *,
        logical_body: str,
    ) -> None:
        """Bind one exact authored fence occurrence to its full visible body."""

        if self._finalized:
            raise DocxSemanticsV3Error("semantic session is already finalized")
        if isinstance(record, FencedSourceIdentityV3):
            identity = fenced_source_identity_from_mapping_v3(fenced_source_mapping_v3(record))
            if identity != record:
                raise DocxSemanticsV3Error("fenced-source identity is not canonically derived")
        else:
            identity = fenced_source_identity_from_mapping_v3(record)
        if identity.source_sha256 != self._source_sha256:
            raise DocxSemanticsV3Error("fenced-source record belongs to a different source document")
        if any(item.identity.tag == identity.tag for item in self._fenced_sources):
            raise DocxSemanticsV3Error("fenced-source tag collision")
        if hashlib.sha256(logical_body.encode()).hexdigest() != identity.body_sha256:
            raise DocxSemanticsV3Error("fenced-source supplied body hash does not match its record")
        write_canonical_fenced_body_v3(paragraph._p, logical_body)
        sdt = wrap_fenced_paragraph_payload(paragraph._p, identity.tag)
        if sdt.getparent() is not paragraph._p:
            raise DocxSemanticsV3Error("fenced-source inline SDT is detached from its paragraph")
        prove_fenced_source_paragraph(paragraph._p, identity)
        self._fenced_sources.append(FencedSourceBindingV3(identity, paragraph._p))

    def bind_heading(self, paragraph: Any, target: dict[str, Any]) -> None:
        """Bind one ID-bearing Heading to its deterministic bookmark and map."""

        if self._finalized:
            raise DocxSemanticsV3Error("semantic session is already finalized")
        if target.get("kind") != "heading" or not target.get("id"):
            raise DocxSemanticsV3Error("Heading binding requires one addressable Heading target")
        identity = derive_target_identity_v3("heading", str(target["id"]))
        self._claim_source_id(identity.source_id)
        wrap_paragraph_content_with_bookmark(
            paragraph,
            identity.bookmark_name,
            self._reserve_bookmark(identity.bookmark_name),
        )
        self._targets.append(TargetBindingV3(identity, (paragraph._p,)))

    def bind_paragraph_anchor(
        self,
        paragraph: Any,
        anchor: dict[str, Any],
        *,
        direct_parent_source_id: str | None = None,
    ) -> None:
        """Register one ordinary paragraph anchor without creating a bookmark."""

        if self._finalized:
            raise DocxSemanticsV3Error("semantic session is already finalized")
        if anchor.get("block_kind") != "paragraph" or not anchor.get("id"):
            raise DocxSemanticsV3Error("the current runtime slice requires an ordinary paragraph anchor")
        identity = derive_anchor_identity_v3("paragraph", str(anchor["id"]))
        if direct_parent_source_id is not None:
            require_source_id(direct_parent_source_id)
        self._claim_source_id(identity.source_id)
        self._anchors.append(AnchorBindingV3(identity, (paragraph._p,), direct_parent_source_id))

    def bind_ordinary_anchor(
        self,
        elements: tuple[Any, ...],
        anchor: dict[str, Any],
        *,
        direct_parent_source_id: str | None = None,
    ) -> None:
        """Bind a closed non-container ordinary block kind without a Word target."""

        if self._finalized:
            raise DocxSemanticsV3Error("semantic session is already finalized")
        block_kind = str(anchor.get("block_kind") or "")
        if block_kind == "fenced_block":
            block_kind = "code_block"
        if block_kind not in {
            "paragraph",
            "image",
            "table",
            "equation",
            "code_block",
            "list",
            "list_item",
            "block_quote",
            "callout",
        }:
            raise DocxSemanticsV3Error("unsupported ordinary-anchor runtime block kind")
        identity = derive_anchor_identity_v3(block_kind, str(anchor.get("id") or ""))
        if direct_parent_source_id is not None:
            require_source_id(direct_parent_source_id)
        self._claim_source_id(identity.source_id)
        self._anchors.append(AnchorBindingV3(identity, elements, direct_parent_source_id))

    def bind_caption(
        self,
        caption: Any,
        object_elements: tuple[Any, ...],
        target: dict[str, Any],
    ) -> None:
        """Bind an addressable caption target to its exact caption+object group."""

        if self._finalized:
            raise DocxSemanticsV3Error("semantic session is already finalized")
        target_id = target.get("id")
        kind = str(target.get("kind") or "")
        if kind not in {"figure", "table", "equation", "code_block"}:
            raise DocxSemanticsV3Error("unsupported caption target kind")
        if not self._caption_style_bindings:
            raise DocxSemanticsV3Error("caption binding requires the complete managed caption-style bindings")
        prove_caption_paragraph_style(caption._p, kind, self._caption_style_bindings)
        cached_number = str(target.get("number") or "")
        if not cached_number:
            raise DocxSemanticsV3Error("caption target has no cached number")
        self._captions.append(
            CaptionBindingV3(
                kind=kind,  # type: ignore[arg-type]
                caption_element=caption._p,
                object_elements=object_elements,
                source_id=str(target_id) if target_id is not None else None,
                title=str(target.get("title") or ""),
                cached_number=cached_number,
            )
        )
        if target_id is None:
            return
        identity = derive_target_identity_v3(kind, str(target_id))  # type: ignore[arg-type]
        self._claim_source_id(identity.source_id)
        bookmark_id = self._reserve_bookmark(identity.bookmark_name)
        self._wrap_caption_seq_with_bookmark(caption, identity.bookmark_name, bookmark_id)
        physical_elements = (
            (*object_elements, caption._p) if identity.kind == "figure" else (caption._p, *object_elements)
        )
        self._targets.append(TargetBindingV3(identity, physical_elements))

    def render_reference(self, paragraph: Any, reference: dict[str, Any]) -> None:
        """Render one supported same-document stable or ID-less soft reference."""

        if self._finalized:
            raise DocxSemanticsV3Error("semantic session is already finalized")
        if reference.get("resolution_status") != "resolved":
            raise DocxSemanticsV3Error("only resolved semantic references can be rendered")
        cached_number = str(reference.get("cached_number") or "")
        if not cached_number:
            raise DocxSemanticsV3Error("semantic reference has no materializable number")
        selector_kind = reference.get("selector_kind")
        if selector_kind == "stable_id":
            target_id = str(reference.get("target_id") or "")
        elif selector_kind == "heading_path" and reference.get("resolved_target_id") is not None:
            target_id = str(reference["resolved_target_id"])
        else:
            target_id = ""
        if target_id:
            target_kind = str(reference.get("resolved_kind") or "")
            target = derive_target_identity_v3(target_kind, target_id)  # type: ignore[arg-type]
            source_range = reference["range"]
            occurrence = derive_reference_occurrence_identity_v3(
                source_sha256=self._source_sha256,
                source_start=int(source_range["start"]),
                source_end=int(source_range["end"]),
                authored_token=str(reference["raw"]),
                resolved_bookmark_name=target.bookmark_name,
                cached_number=cached_number,
            )
            if any(item.tag == occurrence.tag for item in self._reference_occurrences):
                raise DocxSemanticsV3Error("reference-occurrence tag collision")
            paragraph._p.append(
                inline_reference_sdt(
                    occurrence.tag,
                    bookmark_name=target.bookmark_name,
                    cached_number=cached_number,
                    heading_number_only=target.kind == "heading",
                    alias=str(reference["alias"]) if reference.get("alias") is not None else None,
                )
            )
            self._reference_occurrences.append(occurrence)
            self._stable_reference_target_ids.append(target.source_id)
            return

        if selector_kind != "heading_path":
            raise DocxSemanticsV3Error("unsupported soft Heading reference projection")
        source_range = reference["range"]
        identity = derive_soft_reference_identity_v3(
            source_sha256=self._source_sha256,
            source_start=int(source_range["start"]),
            source_end=int(source_range["end"]),
            authored_token=str(reference["raw"]),
            cached_number=cached_number,
        )
        if any(item.tag == identity.tag for item in self._soft_references):
            raise DocxSemanticsV3Error("soft-reference tag collision")
        paragraph._p.append(
            inline_sdt(identity.tag, soft_reference_visible_text(identity.authored_token, cached_number))
        )
        self._soft_references.append(identity)

    def finalize_document(self) -> None:
        """Wrap bound blocks after the template filler has placed them."""

        if self._finalized:
            return
        # An object's ordinary-anchor wrapper is the inner logical object;
        # addressable caption target wrappers then own caption + that object.
        body = self._document.element.body
        ordered, self._anchor_topology_edges = order_ordinary_anchor_bindings(body, self._anchors)
        for binding in ordered:
            owners = tuple(dict.fromkeys(self._direct_body_owner(item) for item in binding.elements))
            wrap_direct_body_group(owners, binding.identity.tag)
        for binding in self._targets:
            elements = tuple(self._direct_body_owner(element) for element in binding.elements)
            unique = tuple(dict.fromkeys(elements))
            wrap_direct_body_group(unique, binding.identity.tag)
        self._finalized = True

    def _direct_body_owner(self, element: Any) -> Any:
        body = self._document.element.body
        owner = element
        while owner.getparent() is not body:
            owner = owner.getparent()
            if owner is None:
                raise DocxSemanticsV3Error("v3 logical group is detached from the direct document body")
        return owner

    def write_package(self, path: str | Path) -> None:
        """Write exact target/anchor and soft-reference custom XML parts."""

        if not self._finalized:
            raise DocxSemanticsV3Error("semantic session must be finalized before package write")
        parts: list[tuple[str, bytes]] = []
        if self._targets or self._anchors:
            parts.append(
                (
                    TARGET_MAP_NAMESPACE,
                    semantic_map_xml(
                        [binding.identity for binding in self._targets],
                        [binding.identity for binding in self._anchors],
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
                (binding.identity for binding in self._fenced_sources),
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
        if self._captions:
            parts.append(
                (
                    CAPTION_STYLE_BINDING_MAP_NAMESPACE,
                    caption_style_binding_map_xml(self._caption_style_bindings),
                )
            )
        if not parts:
            return
        try:
            inject_custom_xml_parts(Path(path), parts)
        except DocxSemanticsV3Error:
            raise
        except (BadZipFile, KeyError, OSError, etree.XMLSyntaxError) as exc:
            raise DocxSemanticsV3Error("failed to write the v3 semantic package projection") from exc

    def prove_package(self, path: str | Path) -> DocxSemanticsV3Recovery:
        """Reopen and authenticate the exact projection before artifact registration."""

        if not self._finalized:
            raise DocxSemanticsV3Error("semantic session must be finalized before package proof")
        package_path = Path(path)
        try:
            from docx import Document
            from docx.opc.exceptions import OpcError

            recovery = DocxSemanticsV3Recovery.load(package_path, Document(str(package_path)))
        except DocxSemanticsV3Error:
            raise
        except (BadZipFile, KeyError, OSError, OpcError, etree.XMLSyntaxError) as exc:
            raise DocxSemanticsV3Error("failed to reopen the v3 semantic package projection") from exc
        expected_targets = tuple(
            sorted((binding.identity for binding in self._targets), key=lambda item: (item.kind, item.source_id))
        )
        expected_anchors = tuple(
            sorted(
                (binding.identity for binding in self._anchors),
                key=lambda item: (item.block_kind, item.source_id),
            )
        )
        expected_soft = tuple(
            sorted(self._soft_references, key=lambda item: (item.source_start, item.source_end, item.tag))
        )
        expected_occurrences = tuple(
            sorted(
                self._reference_occurrences,
                key=lambda item: (item.source_start, item.source_end, item.tag),
            )
        )
        expected_fenced_sources = tuple(
            sorted(
                (binding.identity for binding in self._fenced_sources),
                key=lambda item: (item.source_start, item.source_end, item.tag),
            )
        )
        if (
            recovery.target_identities != expected_targets
            or recovery.anchor_identities != expected_anchors
            or recovery.anchor_topology_edges != self._anchor_topology_edges
            or recovery.soft_reference_identities != expected_soft
            or recovery.reference_occurrence_identities != expected_occurrences
            or recovery.fenced_source_identities != expected_fenced_sources
            or recovery.caption_style_bindings != (self._caption_style_bindings if self._captions else ())
            or recovery.stable_reference_target_ids != tuple(self._stable_reference_target_ids)
            or recovery.caption_signatures
            != tuple(
                (
                    item.kind,
                    item.source_id,
                    item.title,
                    item.cached_number,
                )
                for item in self._captions
            )
        ):
            raise DocxSemanticsV3Error("reopened v3 semantic projection differs from the render session")
        return recovery

    def _claim_source_id(self, source_id: str) -> None:
        if source_id in self._owned_source_ids:
            raise DocxSemanticsV3Error(f"duplicate semantic/anchor source ID: {source_id}")
        self._owned_source_ids.add(source_id)

    def _reserve_bookmark(self, bookmark_name: str) -> str:
        inventory = build_docx_bookmark_inventory(self._document)
        if bookmark_name.casefold() in inventory.used_name_keys:
            raise DocxSemanticsV3Error(f"semantic bookmark conflicts with existing package name: {bookmark_name}")
        used_ids = set(inventory.used_id_keys)
        for candidate in range(BOOKMARK_ID_MAX + 1):
            raw = str(candidate)
            if bookmark_id_key(raw) not in used_ids:
                return raw
        raise DocxSemanticsV3Error("no portable DOCX bookmark IDs remain")

    @staticmethod
    def _wrap_caption_seq_with_bookmark(
        caption: Any,
        bookmark_name: str,
        bookmark_id: str,
    ) -> None:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        children = list(caption._p)
        begin_index = next(
            (
                index
                for index, child in enumerate(children)
                if child.find(f".//{qn('w:fldChar')}[@{qn('w:fldCharType')}='begin']") is not None
            ),
            None,
        )
        if begin_index is None:
            raise DocxSemanticsV3Error("caption target has no SEQ field")
        end_index = next(
            (
                index
                for index in range(begin_index, len(children))
                if children[index].find(f".//{qn('w:fldChar')}[@{qn('w:fldCharType')}='end']") is not None
            ),
            None,
        )
        if end_index is None:
            raise DocxSemanticsV3Error("caption target has an unterminated SEQ field")
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), bookmark_id)
        start.set(qn("w:name"), bookmark_name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), bookmark_id)
        caption._p.insert(begin_index, start)
        caption._p.insert(end_index + 2, end)
