from __future__ import annotations

import gzip
import hashlib
import json
import os
import stat
import tarfile
from pathlib import Path

import pytest
from jsonschema import Draft202012Validator
from scripts.release import linux_archive

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
CONTRACT = ROOT / "release" / "linux-production-manifest.v1.json"
SCHEMA = ROOT / "release" / "linux-production-manifest.v1.schema.json"


def _write_payload(parent: Path, artifact_id: str = "gui-cli") -> tuple[dict[str, object], dict[str, object], Path]:
    contract = linux_archive.read_contract(CONTRACT)
    artifact = linux_archive.select_artifact(contract, artifact_id)
    payload = parent / str(artifact["sourceDirectory"])
    payload.mkdir()
    required_directories = {"_internal", "configs", "models", "templates"}
    if artifact_id == "gui-cli":
        required_directories.add("assets")
    for name in sorted(required_directories):
        (payload / name).mkdir()
    (payload / "_internal" / "runtime.so").write_bytes(b"runtime\x00bytes")
    (payload / "configs" / "default.yaml").write_text("locale: zh-CN\n", encoding="utf-8")
    (payload / "models" / "model.bin").write_bytes(b"model")
    (payload / "templates" / "empty.txt").write_bytes(b"")
    long_directory = payload / "templates" / ("long-" + "x" * 60)
    long_directory.mkdir()
    (long_directory / "unicode-文档.txt").write_text("long pax path\n", encoding="utf-8")
    if artifact_id == "gui-cli":
        (payload / "assets" / "icon.png").write_bytes(b"png")
    for name in ("LICENSE", "LICENSE_THIRD_PARTY.txt", "NOTICE.txt", "README.md"):
        (payload / name).write_text(f"{name}\n", encoding="utf-8")
    for entry_point in artifact["entryPoints"]:
        path = payload / str(entry_point)
        path.write_bytes(f"#!/bin/sh\necho {entry_point}\n".encode())
        path.chmod(0o755)
    return contract, artifact, payload


def _build(parent: Path, payload: Path, artifact: dict[str, object]) -> Path:
    parent.mkdir()
    destination = parent / str(artifact["archiveName"])
    result = linux_archive.build_archive(
        contract_path=CONTRACT,
        artifact_id=str(artifact["artifactId"]),
        payload_root=payload,
        destination=destination,
    )
    assert result["sha256"] == hashlib.sha256(destination.read_bytes()).hexdigest()
    return destination


def test_linux_production_contract_schema_and_two_versioned_assets_are_frozen() -> None:
    schema = json.loads(SCHEMA.read_text(encoding="utf-8"))
    manifest = json.loads(CONTRACT.read_text(encoding="utf-8"))
    Draft202012Validator.check_schema(schema)
    Draft202012Validator(schema).validate(manifest)
    linux_archive.read_contract(CONTRACT)

    assert manifest["target"] == {
        "os": "linux",
        "distribution": "ubuntu",
        "version": "24.04",
        "arch": "x86_64",
        "platformTag": "linux-x64",
    }
    assert [artifact["archiveName"] for artifact in manifest["artifacts"]] == [
        "DocWen-0.9.1-linux-x64.tar.gz",
        "DocWenCLI-0.9.1-linux-x64.tar.gz",
    ]
    assert [artifact["entryPoints"] for artifact in manifest["artifacts"]] == [
        ["DocWen", "DocWenCLI"],
        ["DocWenCLI"],
    ]
    assert manifest["archive"] == linux_archive._EXPECTED_ARCHIVE_POLICY
    assert manifest["symlinks"] == linux_archive._EXPECTED_SYMLINK_POLICY


def test_linux_archives_are_byte_identical_and_normalize_every_header(tmp_path: Path) -> None:
    contract, artifact, payload = _write_payload(tmp_path)
    first = _build(tmp_path / "a", payload, artifact)

    for path in payload.rglob("*"):
        if path.is_file() and path.name not in artifact["entryPoints"]:
            path.chmod(0o600)
    os.utime(payload / "README.md", (1_800_000_000, 1_800_000_000))
    second = _build(tmp_path / "b", payload, artifact)

    assert first.read_bytes() == second.read_bytes()
    raw = first.read_bytes()
    assert raw[:4] == b"\x1f\x8b\x08\x00"
    assert int.from_bytes(raw[4:8], "little") == contract["archive"]["reproducibilityEpoch"]
    assert raw[8] == 2
    assert raw[9] == 255

    with tarfile.open(first, "r:gz") as archive:
        members = archive.getmembers()
        assert [member.name for member in members] == sorted(
            [member.name for member in members], key=lambda value: value.encode("utf-8")
        )
        assert members[0].name == artifact["topLevelDirectory"]
        assert all(member.uid == 0 and member.gid == 0 for member in members)
        assert all(member.uname == "" and member.gname == "" for member in members)
        assert all(member.mtime == contract["archive"]["reproducibilityEpoch"] for member in members)
        assert all(member.issym() is False and member.islnk() is False for member in members)
        assert any(member.pax_headers.get("path") for member in members)
        entry_names = {f"{artifact['topLevelDirectory']}/{name}" for name in artifact["entryPoints"]}
        assert {member.mode for member in members if member.isdir()} == {0o755}
        assert {member.mode for member in members if member.isfile() and member.name in entry_names} == {0o755}
        assert {member.mode for member in members if member.isfile() and member.name not in entry_names} == {0o644}

        root = str(artifact["topLevelDirectory"])
        embedded = json.loads(archive.extractfile(f"{root}/manifest.json").read())
        checksums = archive.extractfile(f"{root}/SHA256SUMS.txt").read().decode("utf-8")
        assert embedded["schema"] == "docwen.linux.payload.v1"
        assert embedded["product"] == {"name": "DocWen", "version": "0.9.1"}
        assert embedded["artifact"]["archiveName"] == artifact["archiveName"]
        assert [row["path"] for row in embedded["files"]] == sorted(
            [row["path"] for row in embedded["files"]], key=lambda value: value.encode("utf-8")
        )
        checksum_paths = [line.split("  ", 1)[1] for line in checksums.splitlines()]
        assert checksum_paths == sorted(checksum_paths, key=lambda value: value.encode("utf-8"))
        assert "manifest.json" in checksum_paths
        assert "SHA256SUMS.txt" not in checksum_paths


def test_cli_only_contract_builds_its_distinct_archive(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path, "cli-only")
    archive_path = _build(tmp_path / "output", payload, artifact)

    with tarfile.open(archive_path, "r:gz") as archive:
        names = archive.getnames()
        assert f"{artifact['topLevelDirectory']}/DocWenCLI" in names
        assert f"{artifact['topLevelDirectory']}/DocWen" not in names


def test_archive_helper_cli_emits_machine_readable_identity(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    _contract, artifact, payload = _write_payload(tmp_path, "cli-only")
    output_parent = tmp_path / "output"
    output_parent.mkdir()
    destination = output_parent / str(artifact["archiveName"])

    assert (
        linux_archive.main(
            [
                "--contract",
                str(CONTRACT),
                "--artifact",
                "cli-only",
                "--payload-root",
                str(payload),
                "--output",
                str(destination),
            ]
        )
        == 0
    )
    result = json.loads(capsys.readouterr().out)
    assert result["artifactId"] == "cli-only"
    assert result["archiveName"] == "DocWenCLI-0.9.1-linux-x64.tar.gz"
    assert result["sha256"] == hashlib.sha256(destination.read_bytes()).hexdigest()


def test_archive_preserves_safe_relative_internal_symlinks(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    long_directory_name = "long-" + "x" * 60
    long_file_name = "resource-" + "y" * 30 + ".so"
    (payload / "templates" / long_directory_name / long_file_name).write_bytes(b"long symlink target")
    long_target = f"../templates/{long_directory_name}/{long_file_name}"
    first = payload / "configs" / "runtime.so"
    second = payload / "configs" / "runtime-current.so"
    long_link = payload / "configs" / "long-resource.so"
    try:
        first.symlink_to("../_internal/runtime.so")
        second.symlink_to("runtime.so")
        long_link.symlink_to(long_target)
    except OSError:
        pytest.skip("symlink creation is unavailable on this host")
    archive_path = _build(tmp_path / "output", payload, artifact)

    with tarfile.open(archive_path, "r:gz") as archive:
        root = str(artifact["topLevelDirectory"])
        first_member = archive.getmember(f"{root}/configs/runtime.so")
        second_member = archive.getmember(f"{root}/configs/runtime-current.so")
        long_member = archive.getmember(f"{root}/configs/long-resource.so")
        assert first_member.issym() and first_member.linkname == "../_internal/runtime.so"
        assert second_member.issym() and second_member.linkname == "runtime.so"
        assert long_member.issym() and long_member.linkname == long_target
        assert long_member.pax_headers["linkpath"] == long_target
        embedded = json.loads(archive.extractfile(f"{root}/manifest.json").read())
        assert embedded["symlinks"] == [
            {"mode": "0777", "path": "configs/long-resource.so", "target": long_target},
            {"mode": "0777", "path": "configs/runtime-current.so", "target": "runtime.so"},
            {"mode": "0777", "path": "configs/runtime.so", "target": "../_internal/runtime.so"},
        ]
        checksums = archive.extractfile(f"{root}/SHA256SUMS.txt").read().decode("utf-8")
        assert "configs/runtime.so" not in checksums
        assert "configs/runtime-current.so" not in checksums


@pytest.mark.parametrize(
    ("target", "error"),
    [
        ("/etc/passwd", "payload_symlink_absolute"),
        ("../../outside", "payload_symlink_escape"),
        ("missing.so", "payload_symlink_dangling"),
        ("../templates", "payload_symlink_directory_target"),
    ],
)
def test_archive_rejects_unsafe_symlink_targets(tmp_path: Path, target: str, error: str) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    link = payload / "configs" / "unsafe.so"
    try:
        link.symlink_to(target)
    except OSError:
        pytest.skip("symlink creation is unavailable on this host")
    output_parent = tmp_path / "output"
    output_parent.mkdir()

    with pytest.raises(linux_archive.LinuxArchiveError, match=error):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=output_parent / str(artifact["archiveName"]),
        )


def test_archive_rejects_symlink_cycles(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    try:
        (payload / "configs" / "a.so").symlink_to("b.so")
        (payload / "configs" / "b.so").symlink_to("a.so")
    except OSError:
        pytest.skip("symlink creation is unavailable on this host")
    output_parent = tmp_path / "output"
    output_parent.mkdir()

    with pytest.raises(linux_archive.LinuxArchiveError, match="payload_symlink_cycle"):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=output_parent / str(artifact["archiveName"]),
        )


def test_archive_rejects_hardlinks(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    hardlink = payload / "configs" / "hardlink.yaml"
    try:
        os.link(payload / "configs" / "default.yaml", hardlink)
    except OSError:
        pytest.skip("hardlink creation is unavailable on this host")
    output_parent = tmp_path / "output"
    output_parent.mkdir()

    with pytest.raises(linux_archive.LinuxArchiveError, match="payload_hardlink_forbidden"):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=output_parent / str(artifact["archiveName"]),
        )


@pytest.mark.skipif(not hasattr(os, "mkfifo"), reason="FIFO creation is unavailable on this host")
def test_archive_rejects_unexpected_file_types(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    os.mkfifo(payload / "configs" / "pipe")
    output_parent = tmp_path / "output"
    output_parent.mkdir()

    with pytest.raises(linux_archive.LinuxArchiveError, match="payload_nonregular_forbidden"):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=output_parent / str(artifact["archiveName"]),
        )


@pytest.mark.parametrize("name", ["manifest.json", "SHA256SUMS.txt"])
def test_archive_rejects_preexisting_generated_metadata(tmp_path: Path, name: str) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    (payload / name).write_text("untrusted\n", encoding="utf-8")
    output_parent = tmp_path / "output"
    output_parent.mkdir()

    with pytest.raises(linux_archive.LinuxArchiveError, match="payload_generated_metadata_collision"):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=output_parent / str(artifact["archiveName"]),
        )


def test_archive_rejects_unexpected_top_level_and_wrong_names(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    (payload / "debug.log").write_text("not a production payload\n", encoding="utf-8")
    output_parent = tmp_path / "output"
    output_parent.mkdir()

    with pytest.raises(linux_archive.LinuxArchiveError, match=r"payload_top_level_files_unexpected:debug\.log"):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=output_parent / str(artifact["archiveName"]),
        )

    with pytest.raises(linux_archive.LinuxArchiveError, match="archive_destination_name_mismatch"):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=output_parent / "DocWen-linux-x64.tar.gz",
        )


def test_archive_publish_never_clobbers_a_racing_destination(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    output_parent = tmp_path / "output"
    output_parent.mkdir()
    destination = output_parent / str(artifact["archiveName"])
    temporary = destination.with_name(f".{destination.name}.tmp")

    def racing_link(_source: Path, target: Path) -> None:
        Path(target).write_bytes(b"other publisher")
        raise FileExistsError(target)

    monkeypatch.setattr(linux_archive.os, "link", racing_link)
    with pytest.raises(linux_archive.LinuxArchiveError, match="archive_destination_exists"):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=destination,
        )
    assert destination.read_bytes() == b"other publisher"
    assert not temporary.exists()


def test_archive_failure_removes_only_its_private_temporary(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    output_parent = tmp_path / "output"
    output_parent.mkdir()
    destination = output_parent / str(artifact["archiveName"])
    temporary = destination.with_name(f".{destination.name}.tmp")

    def reject_header(_path: Path, _policy: dict[str, object]) -> None:
        raise linux_archive.LinuxArchiveError("synthetic_header_rejection")

    monkeypatch.setattr(linux_archive, "_verify_gzip_header", reject_header)
    with pytest.raises(linux_archive.LinuxArchiveError, match="synthetic_header_rejection"):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=destination,
        )
    assert not destination.exists()
    assert not temporary.exists()


def test_archive_does_not_delete_a_racing_unowned_temporary(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    output_parent = tmp_path / "output"
    output_parent.mkdir()
    destination = output_parent / str(artifact["archiveName"])
    temporary = destination.with_name(f".{destination.name}.tmp")

    def racing_write(
        _payload_root: Path,
        target: Path,
        _manifest: dict[str, object],
        _artifact: dict[str, object],
        _tree: linux_archive.CapturedTree,
        _embedded_manifest: bytes,
        _checksums: bytes,
    ) -> None:
        target.write_bytes(b"other process temporary")
        raise FileExistsError(target)

    monkeypatch.setattr(linux_archive, "_write_archive", racing_write)
    with pytest.raises(FileExistsError):
        linux_archive.build_archive(
            contract_path=CONTRACT,
            artifact_id="gui-cli",
            payload_root=payload,
            destination=destination,
        )
    assert temporary.read_bytes() == b"other process temporary"
    assert not destination.exists()


def test_gzip_stream_has_no_filename_and_round_trips(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    archive_path = _build(tmp_path / "output", payload, artifact)

    with gzip.open(archive_path, "rb") as stream:
        assert stream.read(5) == b"DocWe"
    assert archive_path.read_bytes()[3] & 0x08 == 0


def test_normalized_non_executable_modes_do_not_depend_on_source_read_write_modes(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    source = payload / "_internal" / "runtime.so"
    source.chmod(stat.S_IREAD | stat.S_IWRITE)
    archive_path = _build(tmp_path / "output", payload, artifact)

    with tarfile.open(archive_path, "r:gz") as archive:
        member = archive.getmember(f"{artifact['topLevelDirectory']}/_internal/runtime.so")
        assert member.mode == 0o644


@pytest.mark.skipif(os.name != "posix", reason="POSIX executable mode classification is required")
def test_internal_executable_mode_is_preserved_as_normalized_0755(tmp_path: Path) -> None:
    _contract, artifact, payload = _write_payload(tmp_path)
    helper = payload / "_internal" / "helper"
    helper.write_bytes(b"#!/bin/sh\nexit 0\n")
    helper.chmod(0o751)
    archive_path = _build(tmp_path / "output", payload, artifact)

    with tarfile.open(archive_path, "r:gz") as archive:
        root = str(artifact["topLevelDirectory"])
        member = archive.getmember(f"{root}/_internal/helper")
        embedded = json.loads(archive.extractfile(f"{root}/manifest.json").read())
        helper_row = next(row for row in embedded["files"] if row["path"] == "_internal/helper")
        assert member.mode == 0o755
        assert helper_row["mode"] == "0755"
