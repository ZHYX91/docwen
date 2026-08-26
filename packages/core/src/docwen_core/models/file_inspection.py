"""Structured file inspection and admission facts.

The objects in this module are deliberately application-neutral.  A file is
inspected once at an ingress boundary and the resulting fact can then travel
with a :class:`~docwen_core.models.file_ref.FileRef` without another layer
guessing from the filename.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import StrEnum
from types import MappingProxyType
from typing import Any

FILE_INSPECTION_METADATA_KEY = "_docwen_file_inspection"
FILE_ADMISSION_ACCEPTANCE_METADATA_KEY = "_docwen_file_admission_acceptance"

# Stable diagnostic code -> catalogue key.  Core owns this association so
# each presentation adapter cannot silently drift to a different meaning.
# The catalogues and rendering remain the responsibility of CLI / GUI.
FILE_ADMISSION_MESSAGE_KEYS = MappingProxyType(
    {
        "FILE_FORMAT_COMPATIBLE_TEXT": "file_admission.compatible_text",
        "FILE_FORMAT_SAME_FAMILY_MISMATCH": "file_admission.same_family_mismatch",
        "FILE_FORMAT_CROSS_FAMILY_MISMATCH": "file_admission.cross_family_mismatch",
        "FILE_EXTENSION_UNSUPPORTED": "file_admission.unknown_extension",
        "FILE_EMPTY": "file_admission.empty",
        "FILE_CONTAINER_INVALID": "file_admission.container_invalid",
        "FILE_CONTAINER_UNSUPPORTED": "file_admission.container_unsupported",
        "FILE_CONTAINER_UNRECOGNIZED": "file_admission.container_unrecognized",
        "FILE_CONTENT_UNRECOGNIZED": "file_admission.content_unrecognized",
        "FILE_READ_ERROR": "file_admission.read_error",
        "UNSUPPORTED_FORMAT": "file_admission.unsupported_format",
    }
)


def file_admission_message_key(code: str) -> str | None:
    """Return the catalogue key owned by one stable admission code."""

    return FILE_ADMISSION_MESSAGE_KEYS.get(str(code or "").strip().upper())


class DetectionMethod(StrEnum):
    """Evidence used to identify a file's content."""

    SIGNATURE = "signature"
    CONTAINER = "container"
    TEXT_SNIFF = "text_sniff"
    UNKNOWN = "unknown"


class DetectionConfidence(StrEnum):
    """Confidence of the content-derived format."""

    CERTAIN = "certain"
    PROBABLE = "probable"
    UNVERIFIED = "unverified"


class FormatRelation(StrEnum):
    """Relationship between the declared suffix and detected content."""

    EXACT_MATCH = "exact_match"
    EQUIVALENT_ALIAS = "equivalent_alias"
    COMPATIBLE_TEXT = "compatible_text"
    SAME_FAMILY_MISMATCH = "same_family_mismatch"
    CROSS_FAMILY_MISMATCH = "cross_family_mismatch"
    UNRECOGNIZED_EXTENSION = "unrecognized_extension"
    UNVERIFIED = "unverified"


class AdmissionDecision(StrEnum):
    """The action an ingress must take before execution."""

    ALLOW = "allow"
    ALLOW_WITH_WARNING = "allow_with_warning"
    REQUIRE_EXPLICIT_ACCEPTANCE = "require_explicit_acceptance"
    BLOCK = "block"


class StructureStatus(StrEnum):
    """Structural state of the detected file/container."""

    VALID = "valid"
    NOT_APPLICABLE = "not_applicable"
    INVALID = "invalid"
    UNVERIFIED = "unverified"


@dataclass(frozen=True, slots=True)
class ContentDetection:
    """Content-only detection result, independent of the filename."""

    format: str
    method: DetectionMethod
    confidence: DetectionConfidence
    structure_status: StructureStatus = StructureStatus.NOT_APPLICABLE
    detail_code: str = ""
    detail_message: str = ""


@dataclass(frozen=True, slots=True)
class FileInspection:
    """Canonical file inspection and admission decision."""

    file_path: str
    size_bytes: int
    mtime_ns: int
    extension: str
    declared_format: str
    declared_category: str
    detected_format: str
    detected_category: str
    workflow_category: str
    detection_method: DetectionMethod
    confidence: DetectionConfidence
    structure_status: StructureStatus
    relation: FormatRelation
    decision: AdmissionDecision
    declared_supported: bool
    detected_supported: bool
    warning_code: str = ""
    warning_message: str = ""
    reason_code: str = ""
    reason_message: str = ""
    warnings: tuple[dict[str, Any], ...] = field(default_factory=tuple)
    ooxml_signature: dict[str, Any] = field(default_factory=dict)
    device_id: int = 0
    inode: int = 0
    ctime_ns: int = 0
    content_sha256: str = ""

    @property
    def may_execute(self) -> bool:
        """Whether execution is possible without overriding a hard block."""

        return self.decision is not AdmissionDecision.BLOCK

    @property
    def requires_explicit_acceptance(self) -> bool:
        return self.decision is AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE

    def to_dict(self) -> dict[str, Any]:
        """Serialize using stable machine-readable string values."""

        data = asdict(self)
        data["detection_method"] = self.detection_method.value
        data["confidence"] = self.confidence.value
        data["structure_status"] = self.structure_status.value
        data["relation"] = self.relation.value
        data["decision"] = self.decision.value
        data["warnings"] = [dict(item) for item in self.warnings]
        data["ooxml_signature"] = dict(self.ooxml_signature)
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> FileInspection:
        """Deserialize an inspection frozen in ``FileRef.metadata``."""

        return cls(
            file_path=str(data.get("file_path", "")),
            size_bytes=int(data.get("size_bytes", 0)),
            mtime_ns=int(data.get("mtime_ns", 0)),
            extension=str(data.get("extension", "")),
            declared_format=str(data.get("declared_format", "unknown")),
            declared_category=str(data.get("declared_category", "other")),
            detected_format=str(data.get("detected_format", "unknown")),
            detected_category=str(data.get("detected_category", "other")),
            workflow_category=str(data.get("workflow_category", "other")),
            detection_method=DetectionMethod(str(data.get("detection_method", DetectionMethod.UNKNOWN.value))),
            confidence=DetectionConfidence(str(data.get("confidence", DetectionConfidence.UNVERIFIED.value))),
            structure_status=StructureStatus(str(data.get("structure_status", StructureStatus.UNVERIFIED.value))),
            relation=FormatRelation(str(data.get("relation", FormatRelation.UNVERIFIED.value))),
            decision=AdmissionDecision(str(data.get("decision", AdmissionDecision.BLOCK.value))),
            declared_supported=bool(data.get("declared_supported", False)),
            detected_supported=bool(data.get("detected_supported", False)),
            warning_code=str(data.get("warning_code", "")),
            warning_message=str(data.get("warning_message", "")),
            reason_code=str(data.get("reason_code", "")),
            reason_message=str(data.get("reason_message", "")),
            warnings=tuple(dict(item) for item in data.get("warnings", []) if isinstance(item, dict)),
            ooxml_signature=dict(data.get("ooxml_signature", {})),
            device_id=int(data.get("device_id", 0)),
            inode=int(data.get("inode", 0)),
            ctime_ns=int(data.get("ctime_ns", 0)),
            content_sha256=str(data.get("content_sha256", "")),
        )


def make_admission_acceptance(
    inspection: FileInspection,
    *,
    basis: str = "explicit_user_confirmation",
) -> dict[str, Any]:
    """Build the stable acceptance record for one exact inspection."""

    if not inspection.requires_explicit_acceptance:
        raise ValueError("Explicit acceptance is only valid for a confirmation-required inspection.")
    return {
        "accepted": True,
        "basis": basis,
        "relation": inspection.relation.value,
        "file_path": inspection.file_path,
        "size_bytes": inspection.size_bytes,
        "mtime_ns": inspection.mtime_ns,
        "device_id": inspection.device_id,
        "inode": inspection.inode,
        "ctime_ns": inspection.ctime_ns,
        "content_sha256": inspection.content_sha256,
        "declared_format": inspection.declared_format,
        "detected_format": inspection.detected_format,
    }


def admission_is_satisfied(inspection: FileInspection, metadata: dict[str, Any]) -> bool:
    """Return whether a frozen inspection may cross an execution boundary."""

    if inspection.decision is AdmissionDecision.BLOCK:
        return False
    if not inspection.requires_explicit_acceptance:
        return True
    raw = metadata.get(FILE_ADMISSION_ACCEPTANCE_METADATA_KEY)
    if not isinstance(raw, dict):
        return False
    return (
        raw.get("accepted") is True
        and str(raw.get("relation", "")) == inspection.relation.value
        and str(raw.get("file_path", "")) == inspection.file_path
        and int(raw.get("size_bytes", -1)) == inspection.size_bytes
        and int(raw.get("mtime_ns", -1)) == inspection.mtime_ns
        and int(raw.get("device_id", -1)) == inspection.device_id
        and int(raw.get("inode", -1)) == inspection.inode
        and int(raw.get("ctime_ns", -1)) == inspection.ctime_ns
        and str(raw.get("content_sha256", "")) == inspection.content_sha256
        and str(raw.get("declared_format", "")) == inspection.declared_format
        and str(raw.get("detected_format", "")) == inspection.detected_format
    )


__all__ = [
    "FILE_ADMISSION_ACCEPTANCE_METADATA_KEY",
    "FILE_ADMISSION_MESSAGE_KEYS",
    "FILE_INSPECTION_METADATA_KEY",
    "AdmissionDecision",
    "ContentDetection",
    "DetectionConfidence",
    "DetectionMethod",
    "FileInspection",
    "FormatRelation",
    "StructureStatus",
    "admission_is_satisfied",
    "file_admission_message_key",
    "make_admission_acceptance",
]
