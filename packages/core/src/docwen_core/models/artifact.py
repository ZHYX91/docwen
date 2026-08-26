"""ArtifactManifest — describes a single output artifact produced by a plugin."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

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

    def to_dict(self) -> dict[str, Any]:
        return {
            "artifact_id": self.artifact_id,
            "kind": self.kind,
            "staging_path": self.staging_path,
            "suggested_name": self.suggested_name,
            "logical_path": self.logical_path,
            "media_type": self.media_type,
            "metadata": dict(self.metadata),
            "is_primary": self.is_primary,
        }

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
        )
