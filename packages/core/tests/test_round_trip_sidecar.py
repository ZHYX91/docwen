"""Closed, deterministic round-trip sidecar v1 contracts."""

from __future__ import annotations

import json
import stat
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile, ZipInfo

import pytest
from jsonschema import Draft202012Validator

from docwen_core.round_trip_sidecar import (
    AUTHORED_SOURCE_NAME,
    MANIFEST_NAME,
    NEUTRAL_DOCUMENT_NAME,
    NUMBERING_EXPORT_PLAN_NAME,
    ROUND_TRIP_SIDECAR_MEMBERS,
    ROUND_TRIP_SIDECAR_SCHEMA,
    RoundTripSidecarError,
    read_round_trip_sidecar,
    write_round_trip_sidecar,
)

pytestmark = pytest.mark.unit

_FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)


def _payloads() -> dict[str, bytes]:
    return {
        AUTHORED_SOURCE_NAME: b"# Heading\n",
        NEUTRAL_DOCUMENT_NAME: b'{"schema":"docwen.resolved_document.v1"}\n',
        NUMBERING_EXPORT_PLAN_NAME: b'{"schema":"docwen.numbering_export_plan.v1"}\n',
    }


def _write(path: Path, docx: Path) -> None:
    payloads = _payloads()
    write_round_trip_sidecar(
        path,
        docx_path=docx,
        authored_source=payloads[AUTHORED_SOURCE_NAME],
        neutral_document=payloads[NEUTRAL_DOCUMENT_NAME],
        numbering_export_plan=payloads[NUMBERING_EXPORT_PLAN_NAME],
    )


def test_sidecar_is_deterministic_schema_valid_and_authenticates_all_bytes(tmp_path: Path) -> None:
    docx = tmp_path / "document.docx"
    docx.write_bytes(b"docx-package")
    first = tmp_path / "first.docwen"
    second = tmp_path / "second.docwen"
    _write(first, docx)
    _write(second, docx)

    assert first.read_bytes() == second.read_bytes()
    with ZipFile(first) as archive:
        assert tuple(archive.namelist()) == ROUND_TRIP_SIDECAR_MEMBERS
        assert archive.comment == b""
        assert all(
            item.compress_type == ZIP_STORED
            and item.date_time == _FIXED_ZIP_TIME
            and item.extra == b""
            and item.comment == b""
            and item.create_system == 3
            and (item.external_attr >> 16) == stat.S_IFREG | 0o600
            for item in archive.infolist()
        )
        manifest = json.loads(archive.read(MANIFEST_NAME))
    schema_path = Path(__file__).resolve().parents[3] / "contracts/schemas/docwen.round_trip_sidecar.v1.schema.json"
    Draft202012Validator(json.loads(schema_path.read_text(encoding="utf-8"))).validate(manifest)
    assert manifest["schema"] == ROUND_TRIP_SIDECAR_SCHEMA

    recovered = read_round_trip_sidecar(first, docx_path=docx)
    payloads = _payloads()
    assert recovered.authored_source == payloads[AUTHORED_SOURCE_NAME]
    assert recovered.neutral_document == payloads[NEUTRAL_DOCUMENT_NAME]
    assert recovered.numbering_export_plan == payloads[NUMBERING_EXPORT_PLAN_NAME]


def test_sidecar_never_applies_to_modified_docx(tmp_path: Path) -> None:
    docx = tmp_path / "document.docx"
    docx.write_bytes(b"original-package")
    sidecar = tmp_path / "document.docx.docwen"
    _write(sidecar, docx)

    docx.write_bytes(b"edited-package")
    with pytest.raises(RoundTripSidecarError, match="different DOCX") as error:
        read_round_trip_sidecar(sidecar, docx_path=docx)
    assert error.value.code == "docx_mismatch"


def test_sidecar_rejects_payload_tamper(tmp_path: Path) -> None:
    docx = tmp_path / "document.docx"
    docx.write_bytes(b"docx-package")
    valid = tmp_path / "valid.docwen"
    tampered = tmp_path / "tampered.docwen"
    _write(valid, docx)
    with ZipFile(valid) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    members[AUTHORED_SOURCE_NAME] = b"# Changed\n"
    _write_members(tampered, members)

    with pytest.raises(RoundTripSidecarError, match="differs from its manifest") as error:
        read_round_trip_sidecar(tampered, docx_path=docx)
    assert error.value.code == "payload_digest_mismatch"


@pytest.mark.parametrize(
    "names",
    [
        (*ROUND_TRIP_SIDECAR_MEMBERS, "extra.bin"),
        (
            AUTHORED_SOURCE_NAME,
            AUTHORED_SOURCE_NAME,
            NUMBERING_EXPORT_PLAN_NAME,
            MANIFEST_NAME,
        ),
        ("/authored-source.md", NEUTRAL_DOCUMENT_NAME, NUMBERING_EXPORT_PLAN_NAME, MANIFEST_NAME),
        ("../authored-source.md", NEUTRAL_DOCUMENT_NAME, NUMBERING_EXPORT_PLAN_NAME, MANIFEST_NAME),
    ],
)
def test_sidecar_rejects_extra_duplicate_absolute_and_parent_members(
    tmp_path: Path,
    names: tuple[str, ...],
) -> None:
    docx = tmp_path / "document.docx"
    docx.write_bytes(b"docx-package")
    invalid = tmp_path / "invalid.docwen"
    _write_named_members(invalid, names)

    with pytest.raises(RoundTripSidecarError, match="inventory or order"):
        read_round_trip_sidecar(invalid, docx_path=docx)


def test_sidecar_rejects_symbolic_link_member(tmp_path: Path) -> None:
    docx = tmp_path / "document.docx"
    docx.write_bytes(b"docx-package")
    invalid = tmp_path / "link.docwen"
    with ZipFile(invalid, "w", compression=ZIP_STORED) as archive:
        for name in ROUND_TRIP_SIDECAR_MEMBERS:
            info = _member_info(name)
            info.external_attr = (
                (stat.S_IFLNK | 0o777) if name == AUTHORED_SOURCE_NAME else (stat.S_IFREG | 0o600)
            ) << 16
            archive.writestr(info, b"target" if name == AUTHORED_SOURCE_NAME else b"{}\n")

    with pytest.raises(RoundTripSidecarError, match="symbolic links") as error:
        read_round_trip_sidecar(invalid, docx_path=docx)
    assert error.value.code == "member_link_forbidden"


@pytest.mark.parametrize("archive_comment", [False, True])
def test_sidecar_rejects_noncanonical_zip_metadata(tmp_path: Path, archive_comment: bool) -> None:
    docx = tmp_path / "document.docx"
    docx.write_bytes(b"docx-package")
    invalid = tmp_path / "metadata.docwen"
    with ZipFile(invalid, "w", compression=ZIP_STORED) as archive:
        if archive_comment:
            archive.comment = b"unexpected"
        for name in ROUND_TRIP_SIDECAR_MEMBERS:
            info = _member_info(name)
            if not archive_comment and name == AUTHORED_SOURCE_NAME:
                info.date_time = (2026, 1, 1, 0, 0, 0)
            archive.writestr(info, b"{}\n")

    expected_code = "archive_metadata_invalid" if archive_comment else "member_metadata_invalid"
    with pytest.raises(RoundTripSidecarError) as error:
        read_round_trip_sidecar(invalid, docx_path=docx)
    assert error.value.code == expected_code


def test_sidecar_rejects_compression_bomb_ratio(tmp_path: Path) -> None:
    docx = tmp_path / "document.docx"
    docx.write_bytes(b"docx-package")
    invalid = tmp_path / "bomb.docwen"
    with ZipFile(invalid, "w", compression=ZIP_DEFLATED) as archive:
        for name, raw in (
            (AUTHORED_SOURCE_NAME, b"0" * (1024 * 1024)),
            (NEUTRAL_DOCUMENT_NAME, b"{}\n"),
            (NUMBERING_EXPORT_PLAN_NAME, b"{}\n"),
            (MANIFEST_NAME, b"{}\n"),
        ):
            info = _member_info(name)
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, raw)

    with pytest.raises(RoundTripSidecarError, match="compression ratio") as error:
        read_round_trip_sidecar(invalid, docx_path=docx)
    assert error.value.code == "compression_ratio_invalid"


def _write_members(path: Path, members: dict[str, bytes]) -> None:
    with ZipFile(path, "w", compression=ZIP_STORED) as archive:
        for name in ROUND_TRIP_SIDECAR_MEMBERS:
            archive.writestr(_member_info(name), members[name])


def _write_named_members(path: Path, names: tuple[str, ...]) -> None:
    with ZipFile(path, "w", compression=ZIP_STORED) as archive:
        for name in names:
            archive.writestr(_member_info(name), b"{}\n")


def _member_info(name: str) -> ZipInfo:
    info = ZipInfo(name, date_time=_FIXED_ZIP_TIME)
    info.compress_type = ZIP_STORED
    info.create_system = 3
    info.external_attr = (stat.S_IFREG | 0o600) << 16
    return info
