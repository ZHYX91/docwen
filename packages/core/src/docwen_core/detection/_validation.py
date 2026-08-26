"""Canonical file inspection, validation, and admission decisions.

The suffix is treated as a declaration, never as content evidence.  Every
application ingress should call :func:`inspect_file` once and carry the result
forward with the request instead of independently guessing a route.
"""

from __future__ import annotations

import hashlib
from dataclasses import replace
from pathlib import Path
from typing import Any

from docwen_core.detection._sniffing import (
    SUPPORTED_EXTENSION_FORMATS,
    detect_content_format,
)
from docwen_core.detection.ooxml_signature import (
    inspect_ooxml_signature_graph,
    signature_validation_diagnostic,
)
from docwen_core.errors import ValidationError
from docwen_core.formats.categories import get_category, get_media_type
from docwen_core.models.file_inspection import (
    FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
    FILE_INSPECTION_METADATA_KEY,
    AdmissionDecision,
    ContentDetection,
    DetectionConfidence,
    FileInspection,
    FormatRelation,
    StructureStatus,
    admission_is_satisfied,
)
from docwen_core.paths import filesystem_path

_SUPPORTED_EXTENSIONS: frozenset[str] = frozenset(SUPPORTED_EXTENSION_FORMATS)
_SUPPORTED_FORMATS: frozenset[str] = frozenset(SUPPORTED_EXTENSION_FORMATS.values())
_TEXT_WORKFLOW_FORMATS: frozenset[str] = frozenset({"txt", "markdown"})
_ZIP_PACKAGE_FORMATS: frozenset[str] = frozenset({"docx", "xlsx", "pptx", "odt", "ods", "ofd", "xps", "epub"})
_FileIdentity = tuple[int, int, int, int, int]
_HASH_CHUNK_SIZE = 1024 * 1024


def _file_identity(stat_result: Any) -> _FileIdentity:
    """Return the stable identity used by inspection and explicit consent."""

    return (
        int(stat_result.st_dev),
        int(stat_result.st_ino),
        int(stat_result.st_ctime_ns),
        int(stat_result.st_mtime_ns),
        int(stat_result.st_size),
    )


def _content_sha256(path: Path) -> str:
    """Hash the complete content so admission is bound to exact bytes."""

    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(_HASH_CHUNK_SIZE):
            digest.update(chunk)
    return digest.hexdigest()


class FileAdmissionError(ValidationError):
    """A frozen inspection cannot legally cross the execution boundary."""

    def __init__(self, inspection: FileInspection) -> None:
        message = inspection.reason_message or inspection.warning_message or "File admission was denied."
        super().__init__(message)
        self.error_type = admission_error_type(inspection)
        self.details = {"file": inspection.file_path, "admission": inspection.to_dict()}


class FileAdmissionPathError(ValidationError):
    """An input path crosses a link or junction before content inspection."""

    error_type = "input_is_link"

    def __init__(self, path: Path) -> None:
        super().__init__("Input must not be a link or junction.")
        self.details = {"file": str(path)}


def _path_traverses_link_or_junction(path: Path) -> bool:
    current = filesystem_path(path.expanduser())
    while True:
        if current.is_symlink():
            return True
        is_junction = getattr(current, "is_junction", None)
        if callable(is_junction) and is_junction():
            return True
        parent = current.parent
        if parent == current:
            return False
        current = parent


def admission_error_type(inspection: FileInspection) -> str:
    """Map a Core admission reason to the stable public error identifier."""

    return {
        "FILE_EMPTY": "file_empty",
        "FILE_CONTAINER_INVALID": "file_container_invalid",
        "FILE_CONTAINER_UNSUPPORTED": "file_container_unsupported",
        "FILE_CONTAINER_UNRECOGNIZED": "file_container_unrecognized",
        "FILE_CONTENT_UNRECOGNIZED": "file_content_unrecognized",
        "FILE_FORMAT_CONFIRMATION_REQUIRED": "file_format_confirmation_required",
        "UNSUPPORTED_FORMAT": "unsupported_format",
    }.get(inspection.reason_code, "invalid_input")


def _declared_format(extension: str) -> str:
    return SUPPORTED_EXTENSION_FORMATS.get(extension, extension.lstrip(".").lower() or "unknown")


def _workflow_category(detected_format: str, detected_category: str) -> str:
    if detected_format in _TEXT_WORKFLOW_FORMATS:
        return "markdown"
    return detected_category


def _relation(
    *,
    extension: str,
    declared_format: str,
    declared_supported: bool,
    detected_format: str,
    declared_category: str,
    detected_category: str,
) -> FormatRelation:
    if detected_format in {"", "unknown", "zip", "ole"}:
        return FormatRelation.UNVERIFIED
    if not declared_supported:
        return FormatRelation.UNRECOGNIZED_EXTENSION
    if declared_format == detected_format:
        raw_format = extension.lstrip(".").lower()
        return FormatRelation.EXACT_MATCH if raw_format == detected_format else FormatRelation.EQUIVALENT_ALIAS
    if {declared_format, detected_format} <= _TEXT_WORKFLOW_FORMATS:
        return FormatRelation.COMPATIBLE_TEXT
    declared_workflow = _workflow_category(declared_format, declared_category)
    detected_workflow = _workflow_category(detected_format, detected_category)
    if declared_workflow == detected_workflow and detected_workflow != "other":
        return FormatRelation.SAME_FAMILY_MISMATCH
    return FormatRelation.CROSS_FAMILY_MISMATCH


def _warning_payload(
    code: str,
    message: str,
    *,
    declared_format: str,
    detected_format: str,
    relation: FormatRelation,
    decision: AdmissionDecision,
) -> dict[str, object]:
    return {
        "level": "warning",
        "code": code,
        "message": message,
        "details": {
            "declared_format": declared_format,
            "detected_format": detected_format,
            "relation": relation.value,
            "decision": decision.value,
        },
    }


def _decision_for(
    *,
    detection: ContentDetection,
    relation: FormatRelation,
    declared_format: str,
    detected_supported: bool,
) -> tuple[AdmissionDecision, str, str, str, str]:
    """Return decision, warning code/message, and blocking reason code/message."""

    if detection.structure_status is StructureStatus.INVALID:
        code = detection.detail_code or "FILE_CONTAINER_INVALID"
        message = detection.detail_message or "The input container is corrupt or structurally invalid."
        return AdmissionDecision.BLOCK, "", "", code, message

    if detection.format in {"unknown", "ole"} or detection.confidence is DetectionConfidence.UNVERIFIED:
        code = detection.detail_code or "FILE_CONTENT_UNRECOGNIZED"
        message = detection.detail_message or "The input content could not be verified."
        return AdmissionDecision.BLOCK, "", "", code, message

    if detection.format == "zip":
        if declared_format in _ZIP_PACKAGE_FORMATS:
            return (
                AdmissionDecision.BLOCK,
                "",
                "",
                "FILE_CONTAINER_INVALID",
                f"The file is an ordinary ZIP archive, not a valid {declared_format.upper()} package.",
            )
        return (
            AdmissionDecision.BLOCK,
            "",
            "",
            "FILE_CONTAINER_UNSUPPORTED",
            "The ZIP container is not a supported document package.",
        )

    if not detected_supported:
        return (
            AdmissionDecision.BLOCK,
            "",
            "",
            "UNSUPPORTED_FORMAT",
            f"The detected content format ({detection.format}) is not supported.",
        )

    if relation in {FormatRelation.EXACT_MATCH, FormatRelation.EQUIVALENT_ALIAS}:
        return AdmissionDecision.ALLOW, "", "", "", ""

    if relation is FormatRelation.COMPATIBLE_TEXT:
        message = (
            f"File extension ({declared_format}) and detected content ({detection.format}) use the same text workflow."
        )
        return AdmissionDecision.ALLOW_WITH_WARNING, "FILE_FORMAT_COMPATIBLE_TEXT", message, "", ""

    if relation is FormatRelation.SAME_FAMILY_MISMATCH:
        message = f"File extension ({declared_format}) does not match detected content ({detection.format})."
        return AdmissionDecision.ALLOW_WITH_WARNING, "FILE_FORMAT_SAME_FAMILY_MISMATCH", message, "", ""

    if relation is FormatRelation.UNRECOGNIZED_EXTENSION:
        message = (
            f"The filename extension ({declared_format}) is not supported, but the content was detected "
            f"as {detection.format}. Explicit acceptance is required."
        )
        return (
            AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE,
            "FILE_EXTENSION_UNSUPPORTED",
            message,
            "FILE_FORMAT_CONFIRMATION_REQUIRED",
            message,
        )

    message = (
        f"File extension ({declared_format}) declares a different file family than the detected "
        f"content ({detection.format}). Explicit acceptance is required."
    )
    return (
        AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE,
        "FILE_FORMAT_CROSS_FAMILY_MISMATCH",
        message,
        "FILE_FORMAT_CONFIRMATION_REQUIRED",
        message,
    )


def inspect_file(file_path: str) -> FileInspection:
    """Inspect one file and return the canonical admission result."""

    resolved = str(Path(file_path).expanduser().resolve(strict=False))
    public_path = Path(resolved)
    io_path = filesystem_path(public_path)
    if not io_path.is_file():
        raise FileNotFoundError(f"File not found: {file_path}")

    stat_before = io_path.stat()
    size_bytes = stat_before.st_size
    stat_key = _file_identity(stat_before)
    extension = public_path.suffix.lower()
    declared_format = _declared_format(extension)
    declared_category = get_category(declared_format)
    declared_supported = extension in _SUPPORTED_EXTENSIONS
    detection = detect_content_format(str(io_path))
    detected_format = detection.format
    detected_category = get_category(detected_format)
    workflow_category = _workflow_category(detected_format, detected_category)
    detected_supported = detected_format in _SUPPORTED_FORMATS
    content_sha256 = _content_sha256(io_path)
    relation = _relation(
        extension=extension,
        declared_format=declared_format,
        declared_supported=declared_supported,
        detected_format=detected_format,
        declared_category=declared_category,
        detected_category=detected_category,
    )
    decision, warning_code, warning_message, reason_code, reason_message = _decision_for(
        detection=detection,
        relation=relation,
        declared_format=declared_format,
        detected_supported=detected_supported,
    )

    signature_info = inspect_ooxml_signature_graph(str(io_path), actual_format=detected_format)
    stat_after = io_path.stat()
    if _file_identity(stat_after) != stat_key:
        raise OSError(f"File changed while it was being inspected: {file_path}")
    signature_diagnostic = signature_validation_diagnostic(signature_info)
    warnings: list[dict[str, object]] = []
    if warning_code:
        warnings.append(
            _warning_payload(
                warning_code,
                warning_message,
                declared_format=declared_format,
                detected_format=detected_format,
                relation=relation,
                decision=decision,
            )
        )
    if signature_diagnostic is not None:
        warnings.append(signature_diagnostic.to_dict())

    warning_parts: list[str] = []
    for item in warnings:
        message = str(item.get("message", "")).strip()
        if not message:
            continue
        code = str(item.get("code", "")).strip()
        warning_parts.append(f"[{code}] {message}" if code else message)
    combined_warning = " ".join(warning_parts)
    inspection = FileInspection(
        file_path=resolved,
        size_bytes=size_bytes,
        mtime_ns=stat_after.st_mtime_ns,
        extension=extension,
        declared_format=declared_format,
        declared_category=declared_category,
        detected_format=detected_format,
        detected_category=detected_category,
        workflow_category=workflow_category,
        detection_method=detection.method,
        confidence=detection.confidence,
        structure_status=detection.structure_status,
        relation=relation,
        decision=decision,
        declared_supported=declared_supported,
        detected_supported=detected_supported,
        warning_code=warning_code,
        warning_message=combined_warning,
        reason_code=reason_code,
        reason_message=reason_message,
        warnings=tuple(warnings),
        ooxml_signature=signature_info.to_dict(),
        device_id=stat_after.st_dev,
        inode=stat_after.st_ino,
        ctime_ns=stat_after.st_ctime_ns,
        content_sha256=content_sha256,
    )
    return inspection


def has_supported_filename_declaration(file_path: str) -> bool:
    """Return whether the filename declaration is accepted by file pickers.

    This is deliberately a declaration-layer query. It does not inspect the
    file and must never be used as execution-admission evidence.
    """

    extension = Path(file_path).suffix.lower()
    return not extension or extension in _SUPPORTED_EXTENSIONS


def enforce_file_admission(request: Any) -> Any:
    """Enforce current content-derived facts before runtime execution.

    A missing, stale, or path-mismatched inspection is rebuilt from the actual
    input file.  Explicit acceptance is bound to that exact path/stat snapshot,
    so replacing a file after confirmation can never reuse the old consent.
    """

    refs = []
    changed = False
    for ref in request.input_refs:
        if ref.input_role != "source":
            refs.append(ref)
            continue
        metadata = dict(ref.metadata)
        raw = metadata.get(FILE_INSPECTION_METADATA_KEY)
        lexical_path = Path(ref.path)
        if _path_traverses_link_or_junction(lexical_path):
            raise FileAdmissionPathError(lexical_path)
        current_path = str(lexical_path.expanduser().resolve(strict=False))
        inspection = inspect_file(current_path)
        canonical_fact = inspection.to_dict()
        if not isinstance(raw, dict) or raw != canonical_fact:
            metadata[FILE_INSPECTION_METADATA_KEY] = canonical_fact
            metadata.pop(FILE_ADMISSION_ACCEPTANCE_METADATA_KEY, None)
            changed = True

        if not admission_is_satisfied(inspection, metadata):
            raise FileAdmissionError(inspection)
        normalized_category = inspection.workflow_category
        normalized_format = inspection.detected_format
        normalized_media_type = get_media_type(normalized_format)
        if (
            ref.category == normalized_category
            and ref.format == normalized_format
            and ref.media_type == normalized_media_type
            and metadata == ref.metadata
        ):
            refs.append(ref)
            continue
        refs.append(
            replace(
                ref,
                format=normalized_format,
                category=normalized_category,
                media_type=normalized_media_type,
                metadata=metadata,
            )
        )
        changed = True
    return replace(request, input_refs=refs) if changed else request


__all__ = [
    "FileAdmissionError",
    "FileAdmissionPathError",
    "admission_error_type",
    "enforce_file_admission",
    "has_supported_filename_declaration",
    "inspect_file",
]
