"""Consumer-neutral Artifact Bundle v2 data objects and graph validation."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal, NoReturn, cast

from docwen_core.models.document_node import (
    DOCUMENT_NODE_SCHEMA,
    validate_logical_path,
    validate_markdown_node_path,
)

ARTIFACT_BUNDLE_SCHEMA = "docwen.artifact_bundle.v2"
MACHINE_PROTOCOL_SCHEMA = "docwen.machine.v1"

ArtifactKind = Literal["document", "fragment", "resource"]
EntryRole = Literal[
    "primary",
    "supplementary",
    "ocr_page",
    "section",
    "worksheet",
    "image",
    "original",
]
RelationType = Literal["attachment_of", "fragment_of", "resource_of", "derived_from"]
RelationRole = Literal[
    "attachment",
    "ocr_page",
    "ocr_text",
    "section",
    "worksheet",
    "image",
    "original",
    "preview",
    "source",
    "manifest",
]
PageOcrStatus = Literal[
    "success",
    "no_text",
    "input_missing",
    "unavailable",
    "model_missing",
    "initialization_failed",
    "recognition_failed",
]
_PAGE_OCR_STATUSES = frozenset(
    {
        "success",
        "no_text",
        "input_missing",
        "unavailable",
        "model_missing",
        "initialization_failed",
        "recognition_failed",
    }
)


class ArtifactBundleValidationError(ValueError):
    """A cross-field or graph invariant in Artifact Bundle v2 was violated."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(f"{code}: {message}")
        self.code = code


@dataclass(frozen=True, slots=True)
class BundleProducer:
    """Producer identity embedded in every public DocWen bundle."""

    product_version: str
    name: str = "DocWen"
    machine_protocol: str = MACHINE_PROTOCOL_SCHEMA

    def to_dict(self) -> dict[str, str]:
        return {
            "name": self.name,
            "product_version": self.product_version,
            "machine_protocol": self.machine_protocol,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> BundleProducer:
        return cls(
            name=str(data["name"]),
            product_version=str(data["product_version"]),
            machine_protocol=str(data["machine_protocol"]),
        )


@dataclass(frozen=True, slots=True)
class BundleArtifact:
    """One immutable deliverable node in an Artifact Bundle."""

    artifact_id: str
    kind: ArtifactKind
    locator: str
    suggested_name: str
    media_type: str
    size_bytes: int
    sha256: str
    logical_path: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "artifact_id": self.artifact_id,
            "kind": self.kind,
            "locator": self.locator,
            "suggested_name": self.suggested_name,
            "media_type": self.media_type,
            "size_bytes": self.size_bytes,
            "sha256": self.sha256,
            "logical_path": self.logical_path or self.locator,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> BundleArtifact:
        return cls(
            artifact_id=str(data["artifact_id"]),
            kind=cast(ArtifactKind, data["kind"]),
            locator=str(data["locator"]),
            suggested_name=str(data["suggested_name"]),
            media_type=str(data["media_type"]),
            size_bytes=int(data["size_bytes"]),
            sha256=str(data["sha256"]),
            logical_path=str(data["logical_path"]),
        )


@dataclass(frozen=True, slots=True)
class BundleEntry:
    """A root entry and its task-local role in a bundle graph."""

    artifact_id: str
    role: EntryRole
    ordinal: int
    preferred: bool

    def to_dict(self) -> dict[str, Any]:
        return {
            "artifact_id": self.artifact_id,
            "role": self.role,
            "ordinal": self.ordinal,
            "preferred": self.preferred,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> BundleEntry:
        return cls(
            artifact_id=str(data["artifact_id"]),
            role=cast(EntryRole, data["role"]),
            ordinal=int(data["ordinal"]),
            preferred=bool(data["preferred"]),
        )


@dataclass(frozen=True, slots=True)
class PageFragmentSemantics:
    """Closed physical-page facts carried by an OCR fragment relation."""

    page_index: int
    page_count: int
    ocr_status: PageOcrStatus
    source_page: int
    fragment_kind: Literal["page"] = "page"

    def to_dict(self) -> dict[str, Any]:
        return {
            "fragment_kind": self.fragment_kind,
            "page_index": self.page_index,
            "page_count": self.page_count,
            "ocr_status": self.ocr_status,
            "source_page": self.source_page,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PageFragmentSemantics:
        return cls(
            fragment_kind=cast(Literal["page"], data["fragment_kind"]),
            page_index=data["page_index"],
            page_count=data["page_count"],
            ocr_status=cast(PageOcrStatus, data["ocr_status"]),
            source_page=data["source_page"],
        )


@dataclass(frozen=True, slots=True)
class PageResourceSemantics:
    """Proven physical-page origin carried by an exported resource relation."""

    source_page: int

    def to_dict(self) -> dict[str, int]:
        return {"source_page": self.source_page}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PageResourceSemantics:
        return cls(source_page=data["source_page"])


@dataclass(frozen=True, slots=True)
class BundleRelation:
    """A directed semantic relation between two bundle artifacts."""

    type: RelationType
    source_artifact_id: str
    target_artifact_id: str
    role: RelationRole
    ordinal: int | None = None
    page_fragment: PageFragmentSemantics | None = None
    page_resource: PageResourceSemantics | None = None

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "type": self.type,
            "source_artifact_id": self.source_artifact_id,
            "target_artifact_id": self.target_artifact_id,
            "role": self.role,
        }
        if self.ordinal is not None:
            payload["ordinal"] = self.ordinal
        if self.page_fragment is not None:
            payload["page_fragment"] = self.page_fragment.to_dict()
        if self.page_resource is not None:
            payload["page_resource"] = self.page_resource.to_dict()
        return payload

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> BundleRelation:
        return cls(
            type=cast(RelationType, data["type"]),
            source_artifact_id=str(data["source_artifact_id"]),
            target_artifact_id=str(data["target_artifact_id"]),
            role=cast(RelationRole, data["role"]),
            ordinal=int(data["ordinal"]) if "ordinal" in data else None,
            page_fragment=(PageFragmentSemantics.from_dict(data["page_fragment"]) if "page_fragment" in data else None),
            page_resource=(PageResourceSemantics.from_dict(data["page_resource"]) if "page_resource" in data else None),
        )


@dataclass(frozen=True, slots=True)
class ArtifactBundle:
    """A fully committed, integrity-pinned Artifact Bundle v2."""

    bundle_id: str
    task_id: str
    producer: BundleProducer
    artifacts: tuple[BundleArtifact, ...]
    entries: tuple[BundleEntry, ...]
    relations: tuple[BundleRelation, ...] = ()
    schema: str = ARTIFACT_BUNDLE_SCHEMA
    layout_schema: str = "docwen.artifact_layout.v1"

    def to_dict(self) -> dict[str, Any]:
        artifact_payloads = [item.to_dict() for item in self.artifacts]
        payload: dict[str, Any] = {
            "schema": self.schema,
            "bundle_id": self.bundle_id,
            "task_id": self.task_id,
            "producer": self.producer.to_dict(),
            "artifacts": artifact_payloads,
            "entries": [item.to_dict() for item in self.entries],
            "relations": [item.to_dict() for item in self.relations],
        }
        payload["layout_schema"] = self.layout_schema
        return payload

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ArtifactBundle:
        return cls(
            schema=str(data["schema"]),
            bundle_id=str(data["bundle_id"]),
            task_id=str(data["task_id"]),
            producer=BundleProducer.from_dict(data["producer"]),
            artifacts=tuple(BundleArtifact.from_dict(item) for item in data["artifacts"]),
            entries=tuple(BundleEntry.from_dict(item) for item in data["entries"]),
            relations=tuple(BundleRelation.from_dict(item) for item in data["relations"]),
            layout_schema=str(data["layout_schema"]),
        )


@dataclass(frozen=True, slots=True)
class BundleDraftArtifact:
    """A deliverable path plus semantics before integrity commit."""

    artifact_id: str
    kind: ArtifactKind
    path: str
    suggested_name: str
    media_type: str
    logical_path: str | None = None


@dataclass(frozen=True, slots=True)
class BundleDraft:
    """Plugin/route output awaiting runtime path and graph validation."""

    artifacts: tuple[BundleDraftArtifact, ...]
    entries: tuple[BundleEntry, ...]
    relations: tuple[BundleRelation, ...] = ()


def _fail(code: str, message: str) -> NoReturn:
    raise ArtifactBundleValidationError(code, message)


def _validate_relative_locator(locator: str) -> None:
    segments = locator.split("/")
    if (
        "\\" in locator
        or locator.startswith("/")
        or (len(locator) >= 2 and locator[0].isalpha() and locator[1] == ":")
        or any(segment in {"", ".", ".."} for segment in segments)
    ):
        _fail("invalid_artifact_locator", f"artifact locator must be a normalized relative POSIX path: {locator!r}")


def _is_positive_int(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _validate_bundle_structure(
    artifacts: tuple[BundleArtifact, ...] | tuple[BundleDraftArtifact, ...],
    entries: tuple[BundleEntry, ...],
    relations: tuple[BundleRelation, ...],
) -> None:
    """Validate consumer-neutral graph and page semantics without touching paths."""

    artifact_ids = [artifact.artifact_id for artifact in artifacts]
    if not artifacts:
        _fail("empty_artifacts", "a bundle must contain at least one artifact")
    if len(artifact_ids) != len(set(artifact_ids)):
        _fail("duplicate_artifact_id", "artifact_id values must be unique within a bundle")

    artifact_by_id = {artifact.artifact_id: artifact for artifact in artifacts}
    for artifact in artifacts:
        if (
            artifact.suggested_name in {"", ".", ".."}
            or "/" in artifact.suggested_name
            or "\\" in artifact.suggested_name
        ):
            _fail("invalid_suggested_name", "suggested_name must be a safe basename")
        logical_path = artifact.logical_path
        if logical_path:
            try:
                validate_logical_path(logical_path)
            except ValueError as exc:
                _fail("invalid_logical_path", str(exc))

    if not entries:
        _fail("empty_entries", "a bundle must expose at least one entry")
    entry_ids = [entry.artifact_id for entry in entries]
    if len(entry_ids) != len(set(entry_ids)):
        _fail("duplicate_entry", "an artifact may appear in entries at most once")
    entry_ordinals = [entry.ordinal for entry in entries]
    if len(entry_ordinals) != len(set(entry_ordinals)):
        _fail("duplicate_entry_ordinal", "entry ordinal values must be unique")
    for entry in entries:
        artifact = artifact_by_id.get(entry.artifact_id)
        if artifact is None:
            _fail("dangling_entry", f"entry references missing artifact {entry.artifact_id!r}")
        if entry.role == "ocr_page" and artifact.kind != "fragment":
            _fail("incompatible_entry_role", "entry role 'ocr_page' requires a fragment artifact")
        if entry.role == "section" and artifact.kind not in {"document", "fragment"}:
            _fail("incompatible_entry_role", "entry role 'section' requires a document or fragment artifact")
        if entry.role == "image" and artifact.kind != "resource":
            _fail("incompatible_entry_role", "entry role 'image' requires a resource artifact")

    if sum(entry.preferred for entry in entries) > 1:
        _fail("preferred_entry_count", "a bundle may have at most one preferred entry")
    primary_document_ids = {
        entry.artifact_id
        for entry in entries
        if entry.role == "primary" and artifact_by_id[entry.artifact_id].kind == "document"
    }

    structural_relations = frozenset({"attachment_of", "fragment_of", "resource_of"})
    relation_roles = {
        "attachment_of": {"attachment"},
        "fragment_of": {"ocr_page", "ocr_text", "section", "worksheet"},
        "resource_of": {"image", "original", "preview", "worksheet", "manifest"},
        "derived_from": {"source", "original"},
    }
    relation_kinds = {
        "attachment_of": ({"document"}, {"document"}),
        "fragment_of": ({"fragment"}, {"document"}),
        "resource_of": ({"resource"}, {"document", "fragment"}),
        "derived_from": ({"document", "fragment", "resource"}, {"document", "fragment", "resource"}),
    }
    adjacency: dict[str, set[str]] = {artifact_id: set() for artifact_id in artifact_ids}
    directed_edges: dict[str, set[str]] = {artifact_id: set() for artifact_id in artifact_ids}
    structural_owner: dict[str, str] = {}
    ordered_relation_slots: set[tuple[str, str, int]] = set()
    page_relations: list[BundleRelation] = []

    for relation in relations:
        source = relation.source_artifact_id
        target = relation.target_artifact_id
        relation_type = relation.type
        if source not in artifact_by_id or target not in artifact_by_id:
            _fail("dangling_relation", f"relation {source!r} -> {target!r} references a missing artifact")
        if source == target:
            _fail("self_relation", f"artifact {source!r} cannot relate to itself")

        source_kinds, target_kinds = relation_kinds[relation_type]
        if (
            artifact_by_id[source].kind not in source_kinds
            or artifact_by_id[target].kind not in target_kinds
            or relation.role not in relation_roles[relation_type]
        ):
            _fail("incompatible_relation", f"relation {relation_type!r} has incompatible kinds or role")

        if relation_type in structural_relations:
            if source in structural_owner:
                _fail("multiple_structural_owners", f"artifact {source!r} has more than one structural owner")
            structural_owner[source] = target
            if source in entry_ids:
                _fail("owned_entry", f"entry artifact {source!r} cannot also have a structural owner")

        if relation_type in {"attachment_of", "fragment_of"} and relation.ordinal is None:
            _fail("missing_relation_ordinal", f"ordered relation {relation_type!r} requires ordinal")
        is_page_fragment = relation_type == "fragment_of" and relation.role == "ocr_page"
        if relation.ordinal is not None and not is_page_fragment:
            slot = (relation_type, target, relation.ordinal)
            if slot in ordered_relation_slots:
                _fail("duplicate_relation_ordinal", f"duplicate ordinal for {relation_type!r} targeting {target!r}")
            ordered_relation_slots.add(slot)

        adjacency[source].add(target)
        adjacency[target].add(source)
        directed_edges[source].add(target)

    for relation in relations:
        is_page_fragment = relation.type == "fragment_of" and relation.role == "ocr_page"
        if is_page_fragment:
            if relation.page_fragment is None:
                _fail("missing_page_fragment_semantics", "fragment_of/ocr_page requires page_fragment semantics")
            if relation.target_artifact_id not in primary_document_ids:
                _fail("unexpected_page_semantics", "physical page fragments must target a primary document entry")
            page_relations.append(relation)
        elif relation.page_fragment is not None:
            _fail("unexpected_page_semantics", "page_fragment is only valid on fragment_of/ocr_page")

        resource_payload_allowed = relation.type == "resource_of" and relation.role in {"image", "original", "preview"}
        if relation.page_resource is not None and not resource_payload_allowed:
            _fail("unexpected_page_semantics", "page_resource is only valid on image/original/preview resources")

    page_relation_by_artifact: dict[str, BundleRelation] = {}
    pages_by_owner: dict[str, list[BundleRelation]] = {}
    for relation in page_relations:
        page = relation.page_fragment
        assert page is not None
        if (
            page.fragment_kind != "page"
            or any(not _is_positive_int(value) for value in (page.page_index, page.page_count, page.source_page))
            or page.page_index > page.page_count
            or page.source_page > page.page_count
            or not isinstance(page.ocr_status, str)  # pyright: ignore[reportUnnecessaryIsInstance]
            or page.ocr_status not in _PAGE_OCR_STATUSES
        ):
            _fail("invalid_page_range", "physical page numbers and page_count must form positive in-range values")
        if relation.ordinal != page.page_index - 1:
            _fail("page_ordinal_mismatch", "page fragment ordinal must equal page_index - 1")
        page_relation_by_artifact[relation.source_artifact_id] = relation
        pages_by_owner.setdefault(relation.target_artifact_id, []).append(relation)

    for owner, owner_relations in pages_by_owner.items():
        counts = {
            relation.page_fragment.page_count for relation in owner_relations if relation.page_fragment is not None
        }
        if len(counts) != 1:
            _fail("page_count_mismatch", f"page fragments for owner {owner!r} disagree on page_count")
        page_count = next(iter(counts))
        page_indexes = [
            relation.page_fragment.page_index for relation in owner_relations if relation.page_fragment is not None
        ]
        if len(page_indexes) != len(set(page_indexes)):
            _fail("duplicate_page_index", f"page fragments for owner {owner!r} repeat page_index")
        expected = set(range(1, page_count + 1))
        if len(owner_relations) != page_count or set(page_indexes) != expected:
            _fail("incomplete_page_sequence", f"page fragments for owner {owner!r} do not cover 1..page_count")
        source_pages = {
            relation.page_fragment.source_page for relation in owner_relations if relation.page_fragment is not None
        }
        if len(source_pages) != page_count or source_pages != expected:
            _fail("page_source_mismatch", f"source_page values for owner {owner!r} do not cover 1..page_count")

    for relation in relations:
        if relation.type != "resource_of" or relation.role not in {"image", "original", "preview"}:
            continue
        page_resource = relation.page_resource
        if page_resource is not None and not _is_positive_int(page_resource.source_page):
            _fail("invalid_page_range", "page_resource source_page must be a positive integer")
        target_page_relation = page_relation_by_artifact.get(relation.target_artifact_id)
        if target_page_relation is not None:
            target_page = target_page_relation.page_fragment
            assert target_page is not None
            if page_resource is None or page_resource.source_page != target_page.source_page:
                _fail("resource_page_mismatch", "page-owned resource does not match its target page fragment")
            continue
        if artifact_by_id[relation.target_artifact_id].kind == "fragment" and page_resource is not None:
            _fail("resource_page_mismatch", "page_resource targets a fragment without physical-page semantics")
        if page_resource is not None and relation.target_artifact_id not in primary_document_ids:
            _fail("resource_page_mismatch", "document-owned page resource must target a primary document entry")
        if page_resource is not None and relation.target_artifact_id in pages_by_owner:
            _fail("resource_page_mismatch", "proven page resource must target the matching page fragment")

    visited: set[str] = set()

    def visit_for_cycle(artifact_id: str, active: set[str]) -> None:
        if artifact_id in active:
            _fail("relation_cycle", f"relation graph contains a cycle at {artifact_id!r}")
        if artifact_id in visited:
            return
        active.add(artifact_id)
        for target in directed_edges[artifact_id]:
            visit_for_cycle(target, active)
        active.remove(artifact_id)
        visited.add(artifact_id)

    for artifact_id in artifact_ids:
        visit_for_cycle(artifact_id, set())

    reachable = set(entry_ids)
    pending = list(entry_ids)
    while pending:
        artifact_id = pending.pop()
        for neighbor in adjacency[artifact_id] - reachable:
            reachable.add(neighbor)
            pending.append(neighbor)
    orphans = set(artifact_ids) - reachable
    if orphans:
        _fail("orphan_artifact", f"artifacts are not connected to any entry: {sorted(orphans)!r}")


def validate_artifact_bundle_draft(draft: BundleDraft) -> None:
    """Validate a Bundle draft before any staging-root or artifact file I/O."""

    _validate_bundle_structure(draft.artifacts, draft.entries, draft.relations)


def validate_artifact_bundle(bundle: ArtifactBundle) -> None:
    """Validate graph/page facts first, then committed locator invariants."""

    if bundle.schema != ARTIFACT_BUNDLE_SCHEMA:
        _fail("unsupported_bundle_schema", f"unsupported artifact bundle schema: {bundle.schema!r}")
    if bundle.layout_schema not in {"docwen.artifact_layout.v1", DOCUMENT_NODE_SCHEMA}:
        _fail("unsupported_layout_schema", f"unsupported layout schema: {bundle.layout_schema!r}")
    _validate_bundle_structure(bundle.artifacts, bundle.entries, bundle.relations)
    locators = [artifact.locator for artifact in bundle.artifacts]
    if len(locators) != len(set(locators)):
        _fail("duplicate_artifact_locator", "artifact locators must be unique within a bundle")
    for artifact in bundle.artifacts:
        _validate_relative_locator(artifact.locator)
    logical_paths = [artifact.logical_path or artifact.locator for artifact in bundle.artifacts]
    if len(logical_paths) != len(set(logical_paths)):
        _fail("duplicate_logical_path", "artifact logical_path values must be unique within a bundle")
    for artifact, logical_path in zip(bundle.artifacts, logical_paths, strict=True):
        try:
            validate_logical_path(logical_path)
            if bundle.layout_schema == DOCUMENT_NODE_SCHEMA and artifact.media_type == "text/markdown":
                validate_markdown_node_path(logical_path)
        except ValueError as exc:
            _fail("invalid_logical_path", str(exc))


__all__ = [
    "ARTIFACT_BUNDLE_SCHEMA",
    "MACHINE_PROTOCOL_SCHEMA",
    "ArtifactBundle",
    "ArtifactBundleValidationError",
    "BundleArtifact",
    "BundleDraft",
    "BundleDraftArtifact",
    "BundleEntry",
    "BundleProducer",
    "BundleRelation",
    "PageFragmentSemantics",
    "PageOcrStatus",
    "PageResourceSemantics",
    "validate_artifact_bundle",
    "validate_artifact_bundle_draft",
]
