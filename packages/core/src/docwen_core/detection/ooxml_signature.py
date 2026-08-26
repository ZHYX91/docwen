"""Presence-only OOXML digital-signature graph inspection.

This module deliberately inspects OPC package structure only.  It does not
parse or validate XML digital signatures and must never be used as evidence of
document integrity, signer identity, certificate trust, timestamps, or
revocation status.
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import asdict, dataclass, replace
from typing import Any
from xml.etree import ElementTree

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest
from docwen_core.models.result import ConversionDiagnostic

OOXML_SIGNATURE_INFO_METADATA_KEY = "_docwen_ooxml_signature_info"
OOXML_SIGNATURE_VALIDATION_UNAVAILABLE = "OOXML_SIGNATURE_VALIDATION_UNAVAILABLE"
OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED = "OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED"

_OOXML_FORMATS: frozenset[str] = frozenset({"docx", "xlsx", "pptx"})
_ORIGIN_PART = "_xmlsignatures/origin.sigs"
_ORIGIN_RELS_PART = "_xmlsignatures/_rels/origin.sigs.rels"
_SIGNATURE_PREFIX = "_xmlsignatures/"
_ORIGIN_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
_SIGNATURE_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
_ORIGIN_CONTENT_TYPE = "application/vnd.openxmlformats-package.digital-signature-origin"
_SIGNATURE_CONTENT_TYPE = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
_MAX_STRUCTURAL_XML_BYTES = 4 * 1024 * 1024


@dataclass(frozen=True, slots=True)
class OoxmlSignatureInfo:
    """Structural signature-graph fact frozen at request admission."""

    state: str = "not_applicable"
    """One of ``not_applicable``, ``unsigned``, ``complete`` or ``suspicious``."""

    signature_part_count: int = 0
    marker_count: int = 0
    reason: str = ""
    format: str = ""

    @property
    def has_signature_material(self) -> bool:
        return self.state in {"complete", "suspicious"} and self.marker_count > 0

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> OoxmlSignatureInfo:
        state = str(data.get("state", "not_applicable"))
        if state not in {"not_applicable", "unsigned", "complete", "suspicious"}:
            state = "suspicious"
        return cls(
            state=state,
            signature_part_count=max(0, int(data.get("signature_part_count", 0))),
            marker_count=max(0, int(data.get("marker_count", 0))),
            reason=str(data.get("reason", "")),
            format=str(data.get("format", "")),
        )


def inspect_ooxml_signature_graph(
    file_path: str,
    *,
    actual_format: str,
) -> OoxmlSignatureInfo:
    """Classify signature structure from an explicit admitted format fact."""
    if not actual_format.strip():
        raise ValueError("actual_format must be a concrete admitted format")
    normalized_format = actual_format.strip().lower()
    if normalized_format not in _OOXML_FORMATS:
        return OoxmlSignatureInfo(format=normalized_format)

    try:
        with zipfile.ZipFile(file_path) as package:
            normalized_names = [_normalize_part_name(name) for name in package.namelist()]
            names = frozenset(normalized_names)
            protected_names = {
                "[Content_Types].xml",
                "_rels/.rels",
                _ORIGIN_PART,
                _ORIGIN_RELS_PART,
            }
            protected_parts = [
                name
                for name in normalized_names
                if name in protected_names or name.lower().startswith(_SIGNATURE_PREFIX)
            ]
            duplicate_protected_part = len(protected_parts) != len(set(protected_parts))
            signature_parts = sorted(
                name
                for name in names
                if name.startswith(_SIGNATURE_PREFIX)
                and name.lower().endswith(".xml")
                and not name.lower().endswith(".rels")
            )
            root_origin_targets = _relationship_targets(package, "_rels/.rels", _ORIGIN_REL_TYPE)
            signature_targets = _relationship_targets(
                package,
                _ORIGIN_RELS_PART,
                _SIGNATURE_REL_TYPE,
                source_part=_ORIGIN_PART,
            )
            content_types = _content_type_markers(package)
    except (
        OSError,
        KeyError,
        RuntimeError,
        ValueError,
        zipfile.BadZipFile,
        ElementTree.ParseError,
    ):
        return OoxmlSignatureInfo(
            state="suspicious",
            marker_count=0,
            reason="OOXML package or signature graph could not be read structurally",
            format=normalized_format,
        )

    marker_flags = {
        "signature_directory": any(name.lower().startswith(_SIGNATURE_PREFIX) for name in names),
        "origin_part": _ORIGIN_PART in names,
        "origin_relationship": bool(root_origin_targets),
        "origin_relationship_part": _ORIGIN_RELS_PART in names,
        "signature_relationship": bool(signature_targets),
        "signature_part": bool(signature_parts),
        "origin_content_type": content_types["origin"],
        "signature_content_type": content_types["signature"],
    }
    marker_count = sum(marker_flags.values())
    if marker_count == 0:
        return OoxmlSignatureInfo(state="unsigned", format=normalized_format)

    origin_target_ok = root_origin_targets == {_ORIGIN_PART}
    signature_targets_ok = bool(signature_targets) and signature_targets == set(signature_parts)
    complete = (
        all(marker_flags.values())
        and origin_target_ok
        and signature_targets_ok
        and len(signature_parts) == len(signature_targets)
        and not duplicate_protected_part
    )
    if complete:
        return OoxmlSignatureInfo(
            state="complete",
            signature_part_count=len(signature_parts),
            marker_count=marker_count,
            reason="complete structural signature graph detected",
            format=normalized_format,
        )

    missing = [name for name, present in marker_flags.items() if not present]
    if not origin_target_ok:
        missing.append("origin_target_mismatch")
    if not signature_targets_ok:
        missing.append("signature_target_mismatch")
    if duplicate_protected_part:
        missing.append("duplicate_signature_graph_part")
    return OoxmlSignatureInfo(
        state="suspicious",
        signature_part_count=len(signature_parts),
        marker_count=marker_count,
        reason="partial or inconsistent structural signature graph: " + ", ".join(missing),
        format=normalized_format,
    )


def freeze_ooxml_signature_info(request: ConversionRequest) -> ConversionRequest:
    """Return a request whose input refs carry immutable admission facts.

    Existing metadata wins so repeated runtime admission never reclassifies a
    source after the application has frozen it.
    """
    changed = False
    refs: list[FileRef] = []
    for ref in request.input_refs:
        if ref.input_role != "source":
            refs.append(ref)
            continue
        existing = ref.metadata.get(OOXML_SIGNATURE_INFO_METADATA_KEY)
        if isinstance(existing, dict):
            refs.append(ref)
            continue
        info = inspect_ooxml_signature_graph(ref.path, actual_format=ref.format)
        if info.state == "not_applicable":
            refs.append(ref)
            continue
        metadata = dict(ref.metadata)
        metadata[OOXML_SIGNATURE_INFO_METADATA_KEY] = info.to_dict()
        refs.append(replace(ref, metadata=metadata))
        changed = True
    return replace(request, input_refs=refs) if changed else request


def signature_info_for_ref(ref: FileRef) -> OoxmlSignatureInfo:
    """Read a frozen signature fact, falling back to direct shared detection."""
    raw = ref.metadata.get(OOXML_SIGNATURE_INFO_METADATA_KEY)
    if isinstance(raw, dict):
        return OoxmlSignatureInfo.from_dict(raw)
    return inspect_ooxml_signature_graph(ref.path, actual_format=ref.format)


def signature_validation_diagnostic(info: OoxmlSignatureInfo) -> ConversionDiagnostic | None:
    """Build the presence-only limitation warning for signature material."""
    if not info.has_signature_material:
        return None
    graph_label = "complete" if info.state == "complete" else "suspicious or partial"
    return ConversionDiagnostic(
        level="warning",
        code=OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
        message=(
            f"OOXML signature material was detected ({graph_label} structural graph). "
            "DocWen only detected package structure and did not validate document "
            "integrity, signer identity, certificate trust, timestamps, or revocation; "
            "intact and tampered inputs cannot be distinguished."
        ),
    )


def signature_derived_output_diagnostic(
    info: OoxmlSignatureInfo,
) -> ConversionDiagnostic | None:
    """Build the warning attached only after a derived artifact is delivered."""
    if not info.has_signature_material:
        return None
    return ConversionDiagnostic(
        level="warning",
        code=OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED,
        message=(
            "The delivered artifact is derived and unsigned: DocWen did not preserve "
            "or transfer the source OOXML signature. Compare it with the unchanged "
            "source before relying on it."
        ),
    )


def _normalize_part_name(name: str) -> str:
    return name.replace("\\", "/").lstrip("/")


def _relationship_targets(
    package: zipfile.ZipFile,
    rels_part: str,
    relationship_type: str,
    *,
    source_part: str = "",
) -> set[str]:
    normalized_rels = _normalize_part_name(rels_part)
    if normalized_rels not in {_normalize_part_name(name) for name in package.namelist()}:
        return set()
    root = ElementTree.fromstring(_read_structural_xml(package, normalized_rels))
    targets: set[str] = set()
    base = posixpath.dirname(source_part)
    for relationship in root:
        if relationship.attrib.get("Type") != relationship_type:
            continue
        target = relationship.attrib.get("Target", "")
        if not target or relationship.attrib.get("TargetMode", "").lower() == "external":
            continue
        targets.add(_normalize_part_name(posixpath.normpath(posixpath.join(base, target))))
    return targets


def _content_type_markers(package: zipfile.ZipFile) -> dict[str, bool]:
    root = ElementTree.fromstring(_read_structural_xml(package, "[Content_Types].xml"))
    content_types = {node.attrib.get("ContentType", "") for node in root}
    return {
        "origin": _ORIGIN_CONTENT_TYPE in content_types,
        "signature": _SIGNATURE_CONTENT_TYPE in content_types,
    }


def _read_structural_xml(package: zipfile.ZipFile, name: str) -> bytes:
    info = package.getinfo(name)
    if info.file_size > _MAX_STRUCTURAL_XML_BYTES:
        raise ValueError(f"OOXML structural part exceeds inspection limit: {name}")
    return package.read(info)
