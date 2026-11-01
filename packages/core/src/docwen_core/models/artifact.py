"""ArtifactManifest — describes a single output artifact produced by a plugin."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Any, Literal, cast

from docwen_core.models.artifact_bundle import RelationRole


@dataclass(frozen=True, slots=True)
class ArtifactOwnership:
    """Producer-declared structural ownership, available before output placement.

    Each resource occurrence has one owner. Shared bytes may have separate
    artifact identities so each document remains independently portable.
    ``label`` is an optional source-relative child label, not a filesystem path.
    """

    relation_type: Literal["attachment_of", "fragment_of", "resource_of"]
    owner_artifact_id: str
    role: RelationRole
    ordinal: int | None = None
    label: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "relation_type": self.relation_type,
            "owner_artifact_id": self.owner_artifact_id,
            "role": self.role,
            "ordinal": self.ordinal,
            "label": self.label,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ArtifactOwnership:
        return cls(
            relation_type=cast(Literal["attachment_of", "fragment_of", "resource_of"], data["relation_type"]),
            owner_artifact_id=data["owner_artifact_id"],
            role=cast(RelationRole, data["role"]),
            ordinal=data.get("ordinal"),
            label=data.get("label"),
        )


# ── Artifact kind constants ───────────────────────────────────────────

ARTIFACT_KIND_PRIMARY = "primary"
"""The main conversion output."""

ARTIFACT_KIND_AUXILIARY = "auxiliary"
"""A secondary output (e.g. extracted image, intermediate file)."""

ARTIFACT_KIND_IMAGE = "image"
"""An image extracted or generated during conversion."""

ARTIFACT_KIND_LOG = "log"
"""A diagnostic log file."""

ARTIFACT_KIND_MANIFEST = "manifest"
"""A manifest or metadata file describing the conversion."""

ALL_ARTIFACT_KINDS: frozenset[str] = frozenset(
    {
        ARTIFACT_KIND_PRIMARY,
        ARTIFACT_KIND_AUXILIARY,
        ARTIFACT_KIND_IMAGE,
        ARTIFACT_KIND_LOG,
        ARTIFACT_KIND_MANIFEST,
    }
)


@dataclass(slots=True)
class ArtifactManifest:
    """Describes one output artifact produced during conversion.

    Plugins write to staging and return one or more ``ArtifactManifest``
    entries.  The runtime's ``OutputFinalizer`` reads these to perform
    the final placement.

    The *kind* field should use one of the ``ARTIFACT_KIND_*`` constants
    defined in this module.
    """

    artifact_id: str
    """Unique id for this artifact within the task."""

    kind: str
    """Artifact kind.  Use the ``ARTIFACT_KIND_*`` constants:
    ``"primary"``, ``"auxiliary"``, ``"image"``, ``"log"``, ``"manifest"``."""

    staging_path: str
    """Absolute path in the staging directory."""

    suggested_name: str
    """Human-facing fallback filename (without directory)."""

    media_type: str = "application/octet-stream"
    """IANA media type (e.g. ``"text/markdown"``,
    ``"application/vnd.openxmlformats-officedocument.wordprocessingml.document"``)."""

    metadata: dict[str, Any] = field(default_factory=dict)
    """Extra metadata (page count, duration, warnings, etc.).
    A shallow copy is made on serialisation so the caller retains ownership."""

    is_primary: bool = False
    """``True`` if this artifact is the main conversion result.
    Multiple artifacts can be marked primary (e.g. document-split scenarios)."""

    logical_path: str | None = None
    """Portable relative output path assigned by the central layout planner."""

    size_bytes: int | None = None
    """Final byte count, assigned by the runtime after preparing output."""

    sha256: str | None = None
    """SHA-256 of the final bytes, including relocated Markdown links."""

    ownership: ArtifactOwnership | None = None
    """Optional explicit structure shared by layout and Bundle consumers."""

    def to_dict(self) -> dict[str, Any]:
        payload = {
            "artifact_id": self.artifact_id,
            "kind": self.kind,
            "staging_path": self.staging_path,
            "suggested_name": self.suggested_name,
            "logical_path": self.logical_path,
            "media_type": self.media_type,
            "metadata": dict(self.metadata),
            "is_primary": self.is_primary,
            "size_bytes": self.size_bytes,
            "sha256": self.sha256,
        }
        if self.ownership is not None:
            payload["ownership"] = self.ownership.to_dict()
        return payload

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ArtifactManifest:
        return cls(
            artifact_id=data["artifact_id"],
            kind=data["kind"],
            staging_path=data["staging_path"],
            suggested_name=data["suggested_name"],
            logical_path=data.get("logical_path"),
            media_type=data.get("media_type", "application/octet-stream"),
            metadata=dict(data.get("metadata", {})),
            is_primary=data.get("is_primary", False),
            size_bytes=data.get("size_bytes"),
            sha256=data.get("sha256"),
            ownership=ArtifactOwnership.from_dict(data["ownership"]) if data.get("ownership") is not None else None,
        )


def validate_artifact_ownership(artifacts: Sequence[ArtifactManifest]) -> None:
    """Reject invalid producer ownership before publishing any output."""
    by_id = {artifact.artifact_id: artifact for artifact in artifacts}
    if len(by_id) != len(artifacts):
        raise ValueError("duplicate artifact identity")
    roles = {
        "attachment_of": {"attachment"},
        # Physical-page producers retain their richer existing contract; this
        # ownership shorthand does not carry PageFragmentSemantics.
        "fragment_of": {"ocr_text", "section", "worksheet"},
        "resource_of": {"image", "original", "preview", "worksheet", "manifest"},
    }
    slots: set[tuple[str, str, int]] = set()
    for artifact in artifacts:
        ownership = artifact.ownership
        if ownership is None:
            continue
        if (
            not _nonempty_text(ownership.relation_type)
            or ownership.relation_type not in roles
            or not _nonempty_text(ownership.role)
            or ownership.role not in roles[ownership.relation_type]
        ):
            raise ValueError("invalid artifact ownership relation or role")
        if not _nonempty_text(ownership.owner_artifact_id):
            raise ValueError("artifact ownership requires a nonempty owner identity")
        if artifact.is_primary:
            raise ValueError("primary artifact cannot have a structural owner")
        owner = by_id.get(ownership.owner_artifact_id)
        if owner is None:
            raise ValueError("artifact ownership references a missing owner")
        if owner.kind in {"image", "log", "manifest"} or (
            owner.ownership is not None and owner.ownership.relation_type == "resource_of"
        ):
            raise ValueError("artifact owner must be a document or fragment")
        owner_is_fragment = (
            owner.ownership.relation_type == "fragment_of"
            if owner.ownership is not None
            else owner.kind == "auxiliary" and owner.media_type == "text/markdown" and owner.metadata.get("ocr") is True
        )
        if owner.ownership is None and not owner.is_primary and not owner_is_fragment:
            raise ValueError("undeclared auxiliary cannot own document resources")
        if ownership.relation_type != "resource_of" and owner_is_fragment:
            raise ValueError("attachment and fragment owners must be documents, not fragments")
        if ownership.relation_type != "resource_of" and artifact.media_type != "text/markdown":
            raise ValueError("owned document or fragment must be Markdown")
        if ownership.relation_type != "resource_of" and artifact.kind != "auxiliary":
            raise ValueError("owned document or fragment must be an auxiliary artifact")
        if ownership.relation_type == "resource_of" and (
            artifact.kind not in {"auxiliary", "image", "manifest"}
            or (artifact.kind == "manifest" and ownership.role != "manifest")
        ):
            raise ValueError("resource ownership requires an explicit resource artifact")
        if (
            ownership.relation_type == "resource_of"
            and ownership.role in {"image", "preview"}
            and not artifact.media_type.startswith("image/")
        ):
            raise ValueError("image resource ownership requires image media")
        if ownership.ordinal is not None and (type(ownership.ordinal) is not int or ownership.ordinal < 0):
            raise ValueError("artifact ownership ordinal must be a nonnegative integer")
        if ownership.relation_type != "resource_of" and ownership.ordinal is None:
            raise ValueError("owned document or fragment requires an ordinal")
        if ownership.label is not None and not _nonempty_text(ownership.label):
            raise ValueError("artifact ownership label must be nonempty text")
        if ownership.ordinal is not None:
            slot = (ownership.relation_type, ownership.owner_artifact_id, ownership.ordinal)
            if slot in slots:
                raise ValueError("duplicate artifact ownership ordinal")
            slots.add(slot)
        seen = {artifact.artifact_id}
        cursor = owner
        while True:
            if cursor.artifact_id in seen:
                raise ValueError("cyclic artifact ownership")
            seen.add(cursor.artifact_id)
            if cursor.ownership is None:
                break
            next_owner = by_id.get(cursor.ownership.owner_artifact_id)
            if next_owner is None:
                raise ValueError("artifact ownership references a missing owner")
            cursor = next_owner


def _nonempty_text(value: object) -> bool:
    return isinstance(value, str) and bool(value.strip())
