"""Deterministic, authenticated DocWen round-trip sidecar container.

The public sidecar is one regular ``<document>.docx.docwen`` ZIP file so every
DocWen delivery surface can publish it atomically as an Artifact Bundle
resource.  It contains exactly three recovery payloads plus one canonical
manifest binding those payloads to the exact DOCX bytes.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import stat
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, NoReturn
from zipfile import ZIP_DEFLATED, ZIP_STORED, BadZipFile, ZipFile, ZipInfo

from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
)

ROUND_TRIP_SIDECAR_SCHEMA = "docwen.round_trip_sidecar.v1"
ROUND_TRIP_SIDECAR_MEDIA_TYPE = "application/vnd.docwen.round-trip-sidecar+zip"
ROUND_TRIP_SIDECAR_OWNER_METADATA = "round_trip_sidecar_owner_artifact_id"
ROUND_TRIP_SIDECAR_SCHEMA_METADATA = "round_trip_sidecar_schema"
DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

AUTHORED_SOURCE_NAME = "authored-source.md"
NEUTRAL_DOCUMENT_NAME = "neutral-document.json"
NUMBERING_EXPORT_PLAN_NAME = "numbering-export-plan.json"
MANIFEST_NAME = "manifest.json"
ROUND_TRIP_SIDECAR_MEMBERS = (
    AUTHORED_SOURCE_NAME,
    NEUTRAL_DOCUMENT_NAME,
    NUMBERING_EXPORT_PLAN_NAME,
    MANIFEST_NAME,
)

_PAYLOAD_MEDIA_TYPES = {
    AUTHORED_SOURCE_NAME: "text/markdown; charset=utf-8",
    NEUTRAL_DOCUMENT_NAME: RESOLVED_DOCUMENT_MEDIA_TYPE,
    NUMBERING_EXPORT_PLAN_NAME: NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
}
_FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
_HASH_RE = re.compile(r"[0-9a-f]{64}")
_MAX_ARCHIVE_BYTES = 40 * 1024 * 1024
_MAX_MEMBER_BYTES = 16 * 1024 * 1024
_MAX_TOTAL_BYTES = 32 * 1024 * 1024
_MAX_COMPRESSION_RATIO = 100
_HASH_CHUNK_BYTES = 1024 * 1024


class RoundTripSidecarError(ValueError):
    """The sidecar is not the closed v1 container or is not bound to the DOCX."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


@dataclass(frozen=True, slots=True)
class RoundTripSidecarContents:
    """Authenticated recovery bytes from one sidecar."""

    authored_source: bytes
    neutral_document: bytes
    numbering_export_plan: bytes
    manifest: dict[str, Any]


def write_round_trip_sidecar(
    destination: str | Path,
    *,
    docx_path: str | Path,
    authored_source: bytes,
    neutral_document: bytes,
    numbering_export_plan: bytes,
) -> None:
    """Write one deterministic v1 sidecar through a same-directory replace."""

    output = Path(destination)
    docx = Path(docx_path)
    if output.exists() or output.is_symlink():
        _fail("destination_exists", "round-trip sidecar destination must be fresh")
    if docx.is_symlink() or not docx.is_file():
        _fail("docx_invalid", "round-trip sidecar owner must be a regular DOCX file")

    payloads = {
        AUTHORED_SOURCE_NAME: bytes(authored_source),
        NEUTRAL_DOCUMENT_NAME: bytes(neutral_document),
        NUMBERING_EXPORT_PLAN_NAME: bytes(numbering_export_plan),
    }
    _validate_payload_sizes(payloads)
    docx_bytes, docx_sha256 = _file_identity(docx)
    manifest = _manifest_bytes(docx_bytes=docx_bytes, docx_sha256=docx_sha256, payloads=payloads)

    output.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{output.name}.", suffix=".tmp", dir=output.parent)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with ZipFile(temporary, "w", compression=ZIP_STORED, allowZip64=False) as archive:
            for name in ROUND_TRIP_SIDECAR_MEMBERS:
                raw = manifest if name == MANIFEST_NAME else payloads[name]
                archive.writestr(_canonical_member_info(name), raw)
        os.replace(temporary, output)
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise


def read_round_trip_sidecar(
    sidecar_path: str | Path,
    *,
    docx_path: str | Path,
) -> RoundTripSidecarContents:
    """Authenticate one closed v1 container and its exact DOCX owner."""

    sidecar = Path(sidecar_path)
    docx = Path(docx_path)
    if sidecar.is_symlink() or not sidecar.is_file():
        _fail("sidecar_not_regular", "round-trip sidecar is not a regular file")
    if docx.is_symlink() or not docx.is_file():
        _fail("docx_invalid", "round-trip sidecar owner is not a regular DOCX file")
    if sidecar.stat().st_size > _MAX_ARCHIVE_BYTES:
        _fail("archive_too_large", "round-trip sidecar archive exceeds the size limit")

    try:
        with ZipFile(sidecar) as archive:
            if archive.comment:
                _fail("archive_metadata_invalid", "round-trip sidecar archive comment is forbidden")
            infos = archive.infolist()
            _validate_members(infos)
            payloads: dict[str, bytes] = {}
            for info in infos:
                try:
                    raw = archive.read(info)
                except (BadZipFile, RuntimeError, OSError) as exc:
                    raise RoundTripSidecarError(
                        "member_read_failed", "round-trip sidecar member cannot be read"
                    ) from exc
                if len(raw) != info.file_size:
                    _fail("member_size_mismatch", "round-trip sidecar member size changed while reading")
                payloads[info.filename] = raw
    except RoundTripSidecarError:
        raise
    except (BadZipFile, OSError) as exc:
        raise RoundTripSidecarError("archive_invalid", "round-trip sidecar is not a valid ZIP container") from exc

    manifest_raw = payloads.pop(MANIFEST_NAME)
    manifest = _parse_manifest(manifest_raw)
    _validate_payload_sizes(payloads)
    _prove_payload_manifest(manifest, payloads)
    docx_bytes, docx_sha256 = _file_identity(docx)
    docx_record = manifest["docx"]
    if docx_record["bytes"] != docx_bytes or docx_record["sha256"] != docx_sha256:
        _fail("docx_mismatch", "round-trip sidecar belongs to different DOCX bytes")
    return RoundTripSidecarContents(
        authored_source=payloads[AUTHORED_SOURCE_NAME],
        neutral_document=payloads[NEUTRAL_DOCUMENT_NAME],
        numbering_export_plan=payloads[NUMBERING_EXPORT_PLAN_NAME],
        manifest=manifest,
    )


def _manifest_bytes(
    *,
    docx_bytes: int,
    docx_sha256: str,
    payloads: dict[str, bytes],
) -> bytes:
    manifest = {
        "schema": ROUND_TRIP_SIDECAR_SCHEMA,
        "docx": {
            "media_type": DOCX_MEDIA_TYPE,
            "bytes": docx_bytes,
            "sha256": docx_sha256,
        },
        "files": [
            {
                "path": name,
                "media_type": _PAYLOAD_MEDIA_TYPES[name],
                "bytes": len(payloads[name]),
                "sha256": hashlib.sha256(payloads[name]).hexdigest(),
            }
            for name in ROUND_TRIP_SIDECAR_MEMBERS[:-1]
        ],
    }
    return _canonical_json(manifest)


def _parse_manifest(raw: bytes) -> dict[str, Any]:
    try:
        value = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RoundTripSidecarError("manifest_invalid", "round-trip sidecar manifest is not UTF-8 JSON") from exc
    if not isinstance(value, dict) or _canonical_json(value) != raw:
        _fail("manifest_not_canonical", "round-trip sidecar manifest is not canonical JSON")
    if tuple(value) != ("schema", "docx", "files") or value.get("schema") != ROUND_TRIP_SIDECAR_SCHEMA:
        _fail("manifest_schema_invalid", "round-trip sidecar manifest schema is invalid")
    docx = value.get("docx")
    if not isinstance(docx, dict) or tuple(docx) != ("media_type", "bytes", "sha256"):
        _fail("manifest_docx_invalid", "round-trip sidecar DOCX record is invalid")
    if docx.get("media_type") != DOCX_MEDIA_TYPE:
        _fail("manifest_docx_invalid", "round-trip sidecar DOCX media type is invalid")
    _validate_identity(docx, label="DOCX", maximum_bytes=(1 << 63) - 1)
    files = value.get("files")
    if not isinstance(files, list) or len(files) != 3:
        _fail("manifest_files_invalid", "round-trip sidecar must describe exactly three payloads")
    for index, name in enumerate(ROUND_TRIP_SIDECAR_MEMBERS[:-1]):
        record = files[index]
        if not isinstance(record, dict) or tuple(record) != ("path", "media_type", "bytes", "sha256"):
            _fail("manifest_files_invalid", "round-trip sidecar payload record is invalid")
        if record.get("path") != name or record.get("media_type") != _PAYLOAD_MEDIA_TYPES[name]:
            _fail("manifest_files_invalid", "round-trip sidecar payload order or media type is invalid")
        _validate_identity(record, label=name, maximum_bytes=_MAX_MEMBER_BYTES)
    return value


def _validate_members(infos: list[ZipInfo]) -> None:
    names = [info.filename for info in infos]
    if tuple(names) != ROUND_TRIP_SIDECAR_MEMBERS or len(names) != len(set(names)):
        _fail("member_inventory_invalid", "round-trip sidecar member inventory or order is invalid")
    total = 0
    for info in infos:
        path = info.filename
        if (
            info.is_dir()
            or path.startswith(("/", "\\"))
            or "\\" in path
            or any(part in {"", ".", ".."} for part in path.split("/"))
        ):
            _fail("member_path_invalid", "round-trip sidecar member path is unsafe")
        unix_mode = (info.external_attr >> 16) & 0xFFFF
        if stat.S_IFMT(unix_mode) == stat.S_IFLNK:
            _fail("member_link_forbidden", "round-trip sidecar members cannot be symbolic links")
        if (
            info.date_time != _FIXED_ZIP_TIME
            or info.extra
            or info.comment
            or info.create_system != 3
            or unix_mode != stat.S_IFREG | 0o600
        ):
            _fail("member_metadata_invalid", "round-trip sidecar member metadata is not canonical")
        if info.flag_bits & 0x1:
            _fail("member_encrypted", "round-trip sidecar members cannot be encrypted")
        if info.compress_type not in {ZIP_STORED, ZIP_DEFLATED}:
            _fail("member_compression_invalid", "round-trip sidecar compression method is unsupported")
        if info.file_size < 1 or info.file_size > _MAX_MEMBER_BYTES:
            _fail("member_size_invalid", "round-trip sidecar member exceeds the size limit")
        if info.compress_size == 0 or info.file_size > info.compress_size * _MAX_COMPRESSION_RATIO:
            _fail("compression_ratio_invalid", "round-trip sidecar member compression ratio is unsafe")
        total += info.file_size
        if total > _MAX_TOTAL_BYTES:
            _fail("total_size_invalid", "round-trip sidecar expanded size exceeds the limit")


def _validate_payload_sizes(payloads: dict[str, bytes]) -> None:
    if set(payloads) != set(ROUND_TRIP_SIDECAR_MEMBERS[:-1]):
        _fail("payload_inventory_invalid", "round-trip sidecar payload inventory is invalid")
    total = 0
    for name in ROUND_TRIP_SIDECAR_MEMBERS[:-1]:
        size = len(payloads[name])
        if size < 1 or size > _MAX_MEMBER_BYTES:
            _fail("payload_size_invalid", f"round-trip sidecar payload {name!r} exceeds the size limit")
        total += size
    if total > _MAX_TOTAL_BYTES:
        _fail("payload_size_invalid", "round-trip sidecar payload total exceeds the size limit")


def _prove_payload_manifest(manifest: dict[str, Any], payloads: dict[str, bytes]) -> None:
    for record in manifest["files"]:
        raw = payloads[record["path"]]
        if record["bytes"] != len(raw) or record["sha256"] != hashlib.sha256(raw).hexdigest():
            _fail("payload_digest_mismatch", "round-trip sidecar payload differs from its manifest")


def _validate_identity(record: dict[str, Any], *, label: str, maximum_bytes: int) -> None:
    byte_count = record.get("bytes")
    digest = record.get("sha256")
    if (
        not isinstance(byte_count, int)
        or isinstance(byte_count, bool)
        or byte_count < 1
        or byte_count > maximum_bytes
        or not isinstance(digest, str)
        or _HASH_RE.fullmatch(digest) is None
    ):
        _fail("manifest_identity_invalid", f"round-trip sidecar {label} identity is invalid")


def _file_identity(path: Path) -> tuple[int, str]:
    digest = hashlib.sha256()
    byte_count = 0
    try:
        with path.open("rb") as stream:
            while chunk := stream.read(_HASH_CHUNK_BYTES):
                byte_count += len(chunk)
                digest.update(chunk)
    except OSError as exc:
        raise RoundTripSidecarError("file_read_failed", "round-trip sidecar owner cannot be read") from exc
    if byte_count < 1:
        _fail("docx_invalid", "round-trip sidecar owner is empty")
    return byte_count, digest.hexdigest()


def _canonical_json(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, separators=(",", ":")) + "\n").encode("utf-8")


def _canonical_member_info(name: str) -> ZipInfo:
    info = ZipInfo(name, date_time=_FIXED_ZIP_TIME)
    info.compress_type = ZIP_STORED
    info.create_system = 3
    info.external_attr = (stat.S_IFREG | 0o600) << 16
    return info


def _fail(code: str, message: str) -> NoReturn:
    raise RoundTripSidecarError(code, message)


__all__ = [
    "AUTHORED_SOURCE_NAME",
    "DOCX_MEDIA_TYPE",
    "MANIFEST_NAME",
    "NEUTRAL_DOCUMENT_NAME",
    "NUMBERING_EXPORT_PLAN_NAME",
    "ROUND_TRIP_SIDECAR_MEDIA_TYPE",
    "ROUND_TRIP_SIDECAR_MEMBERS",
    "ROUND_TRIP_SIDECAR_OWNER_METADATA",
    "ROUND_TRIP_SIDECAR_SCHEMA",
    "ROUND_TRIP_SIDECAR_SCHEMA_METADATA",
    "RoundTripSidecarContents",
    "RoundTripSidecarError",
    "read_round_trip_sidecar",
    "write_round_trip_sidecar",
]
