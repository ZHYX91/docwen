#!/usr/bin/env python3
"""Create one manifest-bound, deterministic DocWen Linux ``tar.gz``.

This module is deliberately only the archive primitive.  It does not build a
PyInstaller payload, publish an asset, or claim host acceptance.  The caller
must provide one already-built payload directory whose name and top-level
contents satisfy the checked-in Linux production manifest.
"""

from __future__ import annotations

import argparse
import contextlib
import gzip
import hashlib
import io
import json
import os
import re
import stat
import sys
import tarfile
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Any, BinaryIO, cast

STABLE_SEMVER = re.compile(r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")
_CONTRACT_KEYS = {
    "$schema",
    "schemaVersion",
    "manifestId",
    "product",
    "target",
    "archive",
    "symlinks",
    "generatedMetadata",
    "artifacts",
}
_ARTIFACT_KEYS = {
    "artifactId",
    "payloadLayout",
    "sourceDirectory",
    "archiveName",
    "topLevelDirectory",
    "entryPoints",
    "requiredFiles",
    "requiredDirectories",
    "allowedTopLevelFiles",
    "allowedTopLevelDirectories",
}
_EXPECTED_ARCHIVE_POLICY: dict[str, object] = {
    "format": "tar.gz",
    "tarFormat": "pax",
    "entryOrder": "utf8-byte-lexicographic",
    "reproducibilityEpoch": 1704067200,
    "uid": 0,
    "gid": 0,
    "uname": "",
    "gname": "",
    "directoryMode": "0755",
    "entryPointMode": "0755",
    "executableFileMode": "0755",
    "regularFileMode": "0644",
    "symlinkMode": "0777",
    "compressionLevel": 9,
    "gzipHeaderFilename": "",
    "gzipHeaderExtraFlags": 2,
    "gzipHeaderOperatingSystem": 255,
}
_EXPECTED_SYMLINK_POLICY: dict[str, object] = {
    "policy": "relative-internal-files-only",
    "preserve": True,
    "rejectAbsolute": True,
    "rejectEscape": True,
    "rejectDangling": True,
    "rejectCycles": True,
    "rejectDirectoryTargets": True,
}
_EXPECTED_GENERATED_METADATA: dict[str, object] = {
    "payloadManifestPath": "manifest.json",
    "payloadManifestSchema": "docwen.linux.payload.v1",
    "checksumsPath": "SHA256SUMS.txt",
    "checksumsCoverage": "all-regular-files-except-self",
}
_EXPECTED_ARTIFACT_SHAPES = {
    "gui-cli": ("DocWen", "unified-gui-cli-onedir", ("DocWen", "DocWenCLI")),
    "cli-only": ("DocWenCLI", "cli-only-onedir", ("DocWenCLI",)),
}


class LinuxArchiveError(RuntimeError):
    """A stable fail-closed archive contract error."""


@dataclass(frozen=True)
class CapturedFile:
    path: str
    size: int
    sha256: str
    device: int
    inode: int
    mode: int
    modified_ns: int


@dataclass(frozen=True)
class CapturedLink:
    path: str
    target: str
    resolved_target: str


@dataclass(frozen=True)
class CapturedTree:
    directories: tuple[str, ...]
    files: tuple[CapturedFile, ...]
    links: tuple[CapturedLink, ...]


class _DigestingReader:
    def __init__(self, stream: BinaryIO) -> None:
        self._stream = stream
        self._digest = hashlib.sha256()

    def read(self, size: int = -1) -> bytes:
        chunk = self._stream.read(size)
        self._digest.update(chunk)
        return chunk

    def hexdigest(self) -> str:
        return self._digest.hexdigest()


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise LinuxArchiveError(message)


def _require_exact_keys(value: Mapping[str, object], expected: set[str], label: str) -> None:
    _require(set(value) == expected, f"{label}_keys_invalid")


def _as_object(value: object, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise LinuxArchiveError(f"{label}_not_object")
    return cast(dict[str, Any], value)


def _as_string_list(value: object, label: str) -> list[str]:
    if not isinstance(value, list):
        raise LinuxArchiveError(f"{label}_not_array")
    if not all(isinstance(item, str) for item in value):
        raise LinuxArchiveError(f"{label}_item_not_string")
    items = cast(list[str], value)
    _require(len(items) == len(set(items)), f"{label}_duplicate")
    return items


def _utf8_key(value: str) -> bytes:
    try:
        return value.encode("utf-8", errors="strict")
    except UnicodeEncodeError as exc:
        raise LinuxArchiveError(f"path_not_utf8:{value!r}") from exc


def _validate_relative_path(value: str, label: str) -> str:
    _require(value == value.strip(), f"{label}_surrounding_whitespace")
    _require(not any(ord(character) < 32 or ord(character) == 127 for character in value), f"{label}_control")
    _require("\\" not in value, f"{label}_backslash")
    _utf8_key(value)
    path = PurePosixPath(value)
    _require(not path.is_absolute(), f"{label}_absolute")
    _require(value == path.as_posix(), f"{label}_noncanonical")
    _require(bool(path.parts) and all(part not in {"", ".", ".."} for part in path.parts), f"{label}_unsafe")
    return value


def canonical_bytes(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode("utf-8")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validate_contract(manifest: dict[str, Any]) -> None:
    _require_exact_keys(manifest, _CONTRACT_KEYS, "linux_contract")
    _require(manifest["$schema"] == "./linux-production-manifest.v1.schema.json", "linux_contract_schema")
    _require(manifest["schemaVersion"] == 1, "linux_contract_schema_version")
    _require(manifest["manifestId"] == "docwen-linux-production", "linux_contract_manifest_id")

    product = _as_object(manifest["product"], "linux_contract_product")
    _require_exact_keys(product, {"name", "version"}, "linux_contract_product")
    _require(product["name"] == "DocWen", "linux_contract_product_name")
    version = product["version"]
    _require(isinstance(version, str) and STABLE_SEMVER.fullmatch(version) is not None, "linux_contract_version")

    target = _as_object(manifest["target"], "linux_contract_target")
    _require(
        target
        == {
            "os": "linux",
            "distribution": "ubuntu",
            "version": "24.04",
            "arch": "x86_64",
            "platformTag": "linux-x64",
        },
        "linux_contract_target",
    )
    archive = _as_object(manifest["archive"], "linux_contract_archive")
    _require(archive == _EXPECTED_ARCHIVE_POLICY, "linux_contract_archive_policy")
    symlinks = _as_object(manifest["symlinks"], "linux_contract_symlinks")
    _require(symlinks == _EXPECTED_SYMLINK_POLICY, "linux_contract_symlink_policy")
    generated = _as_object(manifest["generatedMetadata"], "linux_contract_generated_metadata")
    _require(generated == _EXPECTED_GENERATED_METADATA, "linux_contract_generated_metadata")

    artifacts = manifest["artifacts"]
    _require(isinstance(artifacts, list) and len(artifacts) == 2, "linux_contract_artifacts")
    _require(
        [item.get("artifactId") if isinstance(item, dict) else None for item in artifacts] == ["gui-cli", "cli-only"],
        "linux_contract_artifact_order",
    )
    for raw_artifact in artifacts:
        artifact = _as_object(raw_artifact, "linux_contract_artifact")
        _require_exact_keys(artifact, _ARTIFACT_KEYS, "linux_contract_artifact")
        artifact_id = artifact["artifactId"]
        _require(
            isinstance(artifact_id, str) and artifact_id in _EXPECTED_ARTIFACT_SHAPES, "linux_contract_artifact_id"
        )
        prefix, layout, expected_entry_points = _EXPECTED_ARTIFACT_SHAPES[artifact_id]
        source_root_name = f"{prefix}_v{version}_linux-x64"
        public_root_name = f"{prefix}-{version}-linux-x64"
        _require(artifact["payloadLayout"] == layout, f"linux_contract_payload_layout:{artifact_id}")
        _require(artifact["sourceDirectory"] == source_root_name, f"linux_contract_source_directory:{artifact_id}")
        _require(artifact["topLevelDirectory"] == public_root_name, f"linux_contract_top_level:{artifact_id}")
        _require(artifact["archiveName"] == f"{public_root_name}.tar.gz", f"linux_contract_archive_name:{artifact_id}")

        entry_points = _as_string_list(artifact["entryPoints"], f"linux_contract_entry_points:{artifact_id}")
        required_files = _as_string_list(artifact["requiredFiles"], f"linux_contract_required_files:{artifact_id}")
        required_directories = _as_string_list(
            artifact["requiredDirectories"], f"linux_contract_required_directories:{artifact_id}"
        )
        allowed_files = _as_string_list(artifact["allowedTopLevelFiles"], f"linux_contract_allowed_files:{artifact_id}")
        allowed_directories = _as_string_list(
            artifact["allowedTopLevelDirectories"], f"linux_contract_allowed_directories:{artifact_id}"
        )
        for label, values in (
            ("entry", entry_points),
            ("required_file", required_files),
            ("required_directory", required_directories),
            ("allowed_file", allowed_files),
            ("allowed_directory", allowed_directories),
        ):
            for value in values:
                _validate_relative_path(value, f"linux_contract_{label}:{artifact_id}")
            _require(values == sorted(values, key=_utf8_key), f"linux_contract_{label}_order:{artifact_id}")
        _require(tuple(entry_points) == expected_entry_points, f"linux_contract_entry_points:{artifact_id}")
        _require(set(entry_points).issubset(required_files), f"linux_contract_entry_points_not_required:{artifact_id}")
        _require(
            {PurePosixPath(path).parts[0] for path in required_files}.issubset(allowed_files),
            f"linux_contract_required_files_not_allowed:{artifact_id}",
        )
        _require(
            {PurePosixPath(path).parts[0] for path in required_directories}.issubset(allowed_directories),
            f"linux_contract_required_directories_not_allowed:{artifact_id}",
        )
        _require(
            not set(allowed_files).intersection(allowed_directories),
            f"linux_contract_allowed_type_collision:{artifact_id}",
        )
        _require(
            not {str(generated["payloadManifestPath"]), str(generated["checksumsPath"])}.intersection(allowed_files),
            f"linux_contract_generated_metadata_in_payload:{artifact_id}",
        )


def read_contract(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise LinuxArchiveError("linux_contract_unreadable") from exc
    manifest = _as_object(value, "linux_contract")
    _validate_contract(manifest)
    return manifest


def select_artifact(manifest: Mapping[str, Any], artifact_id: str) -> dict[str, Any]:
    for value in manifest["artifacts"]:
        if isinstance(value, dict) and value.get("artifactId") == artifact_id:
            return value
    raise LinuxArchiveError(f"linux_contract_artifact_unknown:{artifact_id}")


def _is_reparse_point(status: os.stat_result) -> bool:
    attributes = getattr(status, "st_file_attributes", 0)
    reparse_flag = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)
    return bool(attributes & reparse_flag)


def _identity(status: os.stat_result) -> tuple[int, int, int, int, int]:
    return (
        status.st_dev,
        status.st_ino,
        status.st_size,
        stat.S_IMODE(status.st_mode),
        status.st_mtime_ns,
    )


def _open_flags() -> int:
    flags = os.O_RDONLY
    flags |= getattr(os, "O_BINARY", 0)
    flags |= getattr(os, "O_NOFOLLOW", 0)
    return flags


@contextlib.contextmanager
def _open_regular(path: Path, expected: os.stat_result) -> Iterator[BinaryIO]:
    try:
        descriptor = os.open(path, _open_flags())
    except OSError as exc:
        raise LinuxArchiveError(f"payload_file_open_failed:{path.name}") from exc
    try:
        before = os.fstat(descriptor)
        _require(stat.S_ISREG(before.st_mode), f"payload_nonregular_forbidden:{path.name}")
        _require(_identity(before) == _identity(expected), f"payload_changed_during_capture:{path.name}")
        with os.fdopen(descriptor, "rb", closefd=False) as stream:
            yield stream
        after = os.fstat(descriptor)
        _require(_identity(after) == _identity(before), f"payload_changed_during_read:{path.name}")
    finally:
        os.close(descriptor)


def _hash_regular(path: Path, expected: os.stat_result) -> str:
    digest = hashlib.sha256()
    with _open_regular(path, expected) as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _plain_directory_status(path: Path, label: str) -> os.stat_result:
    try:
        status = os.lstat(path)
    except OSError as exc:
        raise LinuxArchiveError(f"{label}_unavailable") from exc
    _require(stat.S_ISDIR(status.st_mode), f"{label}_not_directory")
    _require(not stat.S_ISLNK(status.st_mode) and not _is_reparse_point(status), f"{label}_link_forbidden")
    return status


def _relative_path(root: Path, path: Path) -> str:
    value = path.relative_to(root).as_posix()
    return _validate_relative_path(value, "payload_path")


def _read_link(path: Path, relative: str) -> CapturedLink:
    try:
        target = os.readlink(path)
    except OSError as exc:
        raise LinuxArchiveError(f"payload_symlink_unreadable:{relative}") from exc
    _require(bool(target), f"payload_symlink_target_empty:{relative}")
    _require("\\" not in target, f"payload_symlink_target_backslash:{relative}")
    _require(
        not any(ord(character) < 32 or ord(character) == 127 for character in target),
        f"payload_symlink_target_control:{relative}",
    )
    _utf8_key(target)
    target_path = PurePosixPath(target)
    _require(not target_path.is_absolute(), f"payload_symlink_absolute:{relative}")
    resolved_parts = list(PurePosixPath(relative).parent.parts)
    for part in target_path.parts:
        if part in {"", "."}:
            continue
        if part == "..":
            _require(bool(resolved_parts), f"payload_symlink_escape:{relative}")
            resolved_parts.pop()
            continue
        resolved_parts.append(part)
    _require(bool(resolved_parts), f"payload_symlink_escape:{relative}")
    resolved_target = "/".join(resolved_parts)
    _validate_relative_path(resolved_target, "payload_symlink_resolved_target")
    return CapturedLink(path=relative, target=target, resolved_target=resolved_target)


def _validate_links(
    links: list[CapturedLink],
    *,
    files: set[str],
    directories: set[str],
) -> None:
    link_map = {link.path: link for link in links}
    for link in links:
        seen = {link.path}
        current = link
        while current.resolved_target in link_map:
            target = current.resolved_target
            _require(target not in seen, f"payload_symlink_cycle:{link.path}")
            seen.add(target)
            current = link_map[target]
        _require(current.resolved_target not in directories, f"payload_symlink_directory_target:{link.path}")
        _require(current.resolved_target in files, f"payload_symlink_dangling:{link.path}")


def _capture_tree(payload_root: Path, artifact: Mapping[str, Any], generated: Mapping[str, Any]) -> CapturedTree:
    root_status = _plain_directory_status(payload_root, "payload_root")
    directories: list[str] = []
    files: list[CapturedFile] = []
    links: list[CapturedLink] = []
    folded: dict[str, str] = {}
    pending = [payload_root]
    while pending:
        current = pending.pop()
        try:
            entries = list(os.scandir(current))
        except OSError as exc:
            relative_current = "." if current == payload_root else _relative_path(payload_root, current)
            raise LinuxArchiveError(f"payload_directory_unreadable:{relative_current}") from exc
        for entry in entries:
            path = Path(entry.path)
            relative = _relative_path(payload_root, path)
            try:
                # ``DirEntry.stat`` on Windows may return a fast-path result
                # with zeroed device/inode/link-count fields.  ``lstat`` keeps
                # identity checks and hard-link rejection effective on both
                # development hosts and the Linux production host.
                status = os.lstat(path)
            except OSError as exc:
                raise LinuxArchiveError(f"payload_entry_unreadable:{relative}") from exc
            _require(status.st_dev == root_status.st_dev, f"payload_mount_forbidden:{relative}")
            _require(
                not _is_reparse_point(status) or stat.S_ISLNK(status.st_mode), f"payload_reparse_forbidden:{relative}"
            )
            casefolded = relative.casefold()
            _require(casefolded not in folded, f"payload_casefold_collision:{folded.get(casefolded)}:{relative}")
            folded[casefolded] = relative
            if stat.S_ISLNK(status.st_mode):
                links.append(_read_link(path, relative))
            elif stat.S_ISDIR(status.st_mode):
                directories.append(relative)
                pending.append(path)
            elif stat.S_ISREG(status.st_mode):
                _require(status.st_nlink == 1, f"payload_hardlink_forbidden:{relative}")
                files.append(
                    CapturedFile(
                        path=relative,
                        size=status.st_size,
                        sha256=_hash_regular(path, status),
                        device=status.st_dev,
                        inode=status.st_ino,
                        mode=stat.S_IMODE(status.st_mode),
                        modified_ns=status.st_mtime_ns,
                    )
                )
            else:
                raise LinuxArchiveError(f"payload_nonregular_forbidden:{relative}")

    directories.sort(key=_utf8_key)
    files.sort(key=lambda item: _utf8_key(item.path))
    links.sort(key=lambda item: _utf8_key(item.path))
    directory_paths = set(directories)
    file_paths = {item.path for item in files}
    link_paths = {item.path for item in links}
    _validate_links(links, files=file_paths, directories=directory_paths)
    paths = directory_paths | file_paths | link_paths
    metadata_paths = {str(generated["payloadManifestPath"]), str(generated["checksumsPath"])}
    collision = sorted(paths.intersection(metadata_paths), key=_utf8_key)
    _require(not collision, f"payload_generated_metadata_collision:{','.join(collision)}")
    missing_files = sorted(set(artifact["requiredFiles"]) - file_paths, key=_utf8_key)
    _require(not missing_files, f"payload_required_files_missing:{','.join(missing_files)}")
    missing_directories = sorted(set(artifact["requiredDirectories"]) - directory_paths, key=_utf8_key)
    _require(not missing_directories, f"payload_required_directories_missing:{','.join(missing_directories)}")
    top_level_files = {path for path in file_paths | link_paths if "/" not in path}
    top_level_directories = {path for path in directory_paths if "/" not in path}
    unexpected_files = sorted(top_level_files - set(artifact["allowedTopLevelFiles"]), key=_utf8_key)
    _require(not unexpected_files, f"payload_top_level_files_unexpected:{','.join(unexpected_files)}")
    unexpected_directories = sorted(top_level_directories - set(artifact["allowedTopLevelDirectories"]), key=_utf8_key)
    _require(
        not unexpected_directories,
        f"payload_top_level_directories_unexpected:{','.join(unexpected_directories)}",
    )
    file_map = {item.path: item for item in files}
    for entry_point in artifact["entryPoints"]:
        _require(entry_point in file_map, f"payload_entry_point_missing:{entry_point}")
        if os.name == "posix":
            _require(file_map[entry_point].mode & 0o111 != 0, f"payload_entry_point_not_executable:{entry_point}")
    return CapturedTree(tuple(directories), tuple(files), tuple(links))


def _captured_status(file: CapturedFile) -> tuple[int, int, int, int, int]:
    return (file.device, file.inode, file.size, file.mode, file.modified_ns)


def _normalized_file_mode(file: CapturedFile, artifact: Mapping[str, Any], policy: Mapping[str, Any]) -> str:
    if file.path in artifact["entryPoints"]:
        return str(policy["entryPointMode"])
    if file.mode & 0o111:
        return str(policy["executableFileMode"])
    return str(policy["regularFileMode"])


def _payload_manifest(
    manifest: Mapping[str, Any], artifact: Mapping[str, Any], tree: CapturedTree
) -> dict[str, object]:
    archive_policy = manifest["archive"]
    return {
        "schema": manifest["generatedMetadata"]["payloadManifestSchema"],
        "product": dict(manifest["product"]),
        "target": dict(manifest["target"]),
        "artifact": {
            "artifactId": artifact["artifactId"],
            "payloadLayout": artifact["payloadLayout"],
            "archiveName": artifact["archiveName"],
            "topLevelDirectory": artifact["topLevelDirectory"],
            "entryPoints": list(artifact["entryPoints"]),
        },
        "normalization": {
            "entryOrder": archive_policy["entryOrder"],
            "reproducibilityEpoch": archive_policy["reproducibilityEpoch"],
            "uid": archive_policy["uid"],
            "gid": archive_policy["gid"],
            "uname": archive_policy["uname"],
            "gname": archive_policy["gname"],
            "directoryMode": archive_policy["directoryMode"],
            "entryPointMode": archive_policy["entryPointMode"],
            "executableFileMode": archive_policy["executableFileMode"],
            "regularFileMode": archive_policy["regularFileMode"],
            "symlinkMode": archive_policy["symlinkMode"],
        },
        "directories": list(tree.directories),
        "files": [
            {
                "path": item.path,
                "bytes": item.size,
                "sha256": item.sha256,
                "mode": _normalized_file_mode(item, artifact, archive_policy),
            }
            for item in tree.files
        ],
        "symlinks": [
            {
                "path": item.path,
                "target": item.target,
                "mode": archive_policy["symlinkMode"],
            }
            for item in tree.links
        ],
    }


def _checksums(tree: CapturedTree, manifest_path: str, manifest_bytes: bytes) -> bytes:
    rows = [(item.path, item.sha256) for item in tree.files]
    rows.append((manifest_path, hashlib.sha256(manifest_bytes).hexdigest()))
    rows.sort(key=lambda item: _utf8_key(item[0]))
    return "".join(f"{digest}  {path}\n" for path, digest in rows).encode("utf-8")


def _tar_info(name: str, *, mode: int, epoch: int, size: int = 0, directory: bool = False) -> tarfile.TarInfo:
    info = tarfile.TarInfo(name)
    info.type = tarfile.DIRTYPE if directory else tarfile.REGTYPE
    info.mode = mode
    info.uid = 0
    info.gid = 0
    info.uname = ""
    info.gname = ""
    info.mtime = epoch
    info.size = 0 if directory else size
    info.pax_headers = {}
    return info


def _tar_link_info(name: str, target: str, *, mode: int, epoch: int) -> tarfile.TarInfo:
    info = _tar_info(name, mode=mode, epoch=epoch)
    info.type = tarfile.SYMTYPE
    info.size = 0
    info.linkname = target
    return info


def _source_status(path: Path, captured: CapturedFile) -> os.stat_result:
    try:
        status = os.lstat(path)
    except OSError as exc:
        raise LinuxArchiveError(f"payload_file_unavailable:{captured.path}") from exc
    _require(
        not stat.S_ISLNK(status.st_mode) and not _is_reparse_point(status), f"payload_link_forbidden:{captured.path}"
    )
    _require(_identity(status) == _captured_status(captured), f"payload_changed_before_archive:{captured.path}")
    return status


def _write_archive(
    payload_root: Path,
    destination: Path,
    manifest: Mapping[str, Any],
    artifact: Mapping[str, Any],
    tree: CapturedTree,
    embedded_manifest: bytes,
    checksums: bytes,
) -> None:
    policy = manifest["archive"]
    generated = manifest["generatedMetadata"]
    epoch = int(policy["reproducibilityEpoch"])
    directory_mode = int(str(policy["directoryMode"]), 8)
    regular_mode = int(str(policy["regularFileMode"]), 8)
    entry_mode = int(str(policy["entryPointMode"]), 8)
    executable_mode = int(str(policy["executableFileMode"]), 8)
    symlink_mode = int(str(policy["symlinkMode"]), 8)
    root_name = str(artifact["topLevelDirectory"])
    metadata = {
        str(generated["payloadManifestPath"]): embedded_manifest,
        str(generated["checksumsPath"]): checksums,
    }
    entries: list[tuple[str, str, object]] = []
    entries.extend((path, "directory", path) for path in tree.directories)
    entries.extend((item.path, "file", item) for item in tree.files)
    entries.extend((item.path, "symlink", item) for item in tree.links)
    entries.extend((path, "metadata", payload) for path, payload in metadata.items())
    entries.sort(key=lambda item: _utf8_key(item[0]))

    opened = False
    try:
        with destination.open("xb") as raw:
            opened = True
            with (
                gzip.GzipFile(
                    filename=str(policy["gzipHeaderFilename"]),
                    mode="wb",
                    compresslevel=int(policy["compressionLevel"]),
                    fileobj=raw,
                    mtime=epoch,
                ) as compressed,
                tarfile.open(fileobj=compressed, mode="w|", format=tarfile.PAX_FORMAT) as archive,
            ):
                archive.addfile(_tar_info(root_name, mode=directory_mode, epoch=epoch, directory=True))
                for relative, kind, value in entries:
                    name = f"{root_name}/{relative}"
                    if kind == "directory":
                        archive.addfile(_tar_info(name, mode=directory_mode, epoch=epoch, directory=True))
                        continue
                    if kind == "metadata":
                        if not isinstance(value, bytes):
                            raise LinuxArchiveError(f"generated_metadata_invalid:{relative}")
                        payload = value
                        archive.addfile(
                            _tar_info(name, mode=regular_mode, epoch=epoch, size=len(payload)), io.BytesIO(payload)
                        )
                        continue
                    if kind == "symlink":
                        if not isinstance(value, CapturedLink):
                            raise LinuxArchiveError(f"captured_link_invalid:{relative}")
                        observed_link = _read_link(payload_root / PurePosixPath(value.path), value.path)
                        _require(observed_link == value, f"payload_symlink_changed_during_archive:{relative}")
                        archive.addfile(_tar_link_info(name, value.target, mode=symlink_mode, epoch=epoch))
                        continue
                    if not isinstance(value, CapturedFile):
                        raise LinuxArchiveError(f"captured_file_invalid:{relative}")
                    captured = value
                    source = payload_root / PurePosixPath(captured.path)
                    expected_status = _source_status(source, captured)
                    if captured.path in artifact["entryPoints"]:
                        mode = entry_mode
                    elif captured.mode & 0o111:
                        mode = executable_mode
                    else:
                        mode = regular_mode
                    with _open_regular(source, expected_status) as source_stream:
                        reader = _DigestingReader(source_stream)
                        archive.addfile(_tar_info(name, mode=mode, epoch=epoch, size=captured.size), reader)
                        _require(
                            reader.hexdigest() == captured.sha256,
                            f"payload_changed_during_archive:{captured.path}",
                        )
            raw.flush()
            os.fsync(raw.fileno())
    except Exception:
        if opened:
            with contextlib.suppress(OSError):
                destination.unlink()
        raise


def _verify_gzip_header(path: Path, policy: Mapping[str, Any]) -> None:
    with path.open("rb") as stream:
        header = stream.read(10)
    _require(len(header) == 10 and header[:3] == b"\x1f\x8b\x08", "archive_gzip_header_invalid")
    _require(header[3] == 0, "archive_gzip_optional_header_forbidden")
    _require(int.from_bytes(header[4:8], "little") == policy["reproducibilityEpoch"], "archive_gzip_mtime_drift")
    _require(header[8] == policy["gzipHeaderExtraFlags"], "archive_gzip_extra_flags_drift")
    _require(header[9] == policy["gzipHeaderOperatingSystem"], "archive_gzip_operating_system_drift")


def _verify_archive(
    path: Path,
    manifest: Mapping[str, Any],
    artifact: Mapping[str, Any],
    tree: CapturedTree,
    embedded_manifest: bytes,
    checksums: bytes,
) -> None:
    policy = manifest["archive"]
    generated = manifest["generatedMetadata"]
    root_name = str(artifact["topLevelDirectory"])
    directory_mode = int(str(policy["directoryMode"]), 8)
    regular_mode = int(str(policy["regularFileMode"]), 8)
    entry_mode = int(str(policy["entryPointMode"]), 8)
    executable_mode = int(str(policy["executableFileMode"]), 8)
    symlink_mode = int(str(policy["symlinkMode"]), 8)
    metadata = {
        str(generated["payloadManifestPath"]): embedded_manifest,
        str(generated["checksumsPath"]): checksums,
    }
    expected: list[tuple[str, str, str | None, int, str | None]] = [
        (root_name, "directory", None, directory_mode, None)
    ]
    logical_entries: list[tuple[str, str, str | None, int, str | None]] = []
    logical_entries.extend((directory, "directory", None, directory_mode, None) for directory in tree.directories)
    logical_entries.extend(
        (
            item.path,
            "file",
            item.sha256,
            (
                entry_mode
                if item.path in artifact["entryPoints"]
                else executable_mode
                if item.mode & 0o111
                else regular_mode
            ),
            None,
        )
        for item in tree.files
    )
    logical_entries.extend((item.path, "symlink", None, symlink_mode, item.target) for item in tree.links)
    logical_entries.extend(
        (relative, "metadata", hashlib.sha256(payload).hexdigest(), regular_mode, None)
        for relative, payload in metadata.items()
    )
    logical_entries.sort(key=lambda item: _utf8_key(item[0]))
    expected.extend(
        (f"{root_name}/{relative}", kind, digest, mode, target)
        for relative, kind, digest, mode, target in logical_entries
    )

    try:
        with tarfile.open(path, "r:gz") as archive:
            members = archive.getmembers()
            _require([member.name for member in members] == [item[0] for item in expected], "archive_entry_order_drift")
            for member, (name, kind, digest, mode, link_target) in zip(members, expected, strict=True):
                _require(member.name == name, f"archive_entry_name_drift:{name}")
                _require(member.uid == 0 and member.gid == 0, f"archive_owner_drift:{name}")
                _require(member.uname == "" and member.gname == "", f"archive_owner_name_drift:{name}")
                _require(member.mtime == policy["reproducibilityEpoch"], f"archive_mtime_drift:{name}")
                _require(member.mode == mode, f"archive_mode_drift:{name}")
                _require(not member.islnk(), f"archive_hardlink_forbidden:{name}")
                allowed_pax_headers = {"path", "linkpath"} if kind == "symlink" else {"path"}
                _require(set(member.pax_headers).issubset(allowed_pax_headers), f"archive_pax_header_unexpected:{name}")
                if "path" in member.pax_headers:
                    pax_path = (
                        member.pax_headers["path"].removesuffix("/") if member.isdir() else member.pax_headers["path"]
                    )
                    _require(pax_path == name, f"archive_pax_path_drift:{name}")
                if "linkpath" in member.pax_headers:
                    _require(member.pax_headers["linkpath"] == link_target, f"archive_pax_linkpath_drift:{name}")
                if kind == "directory":
                    _require(
                        member.isdir() and not member.issym() and member.linkname == "" and member.size == 0,
                        f"archive_directory_invalid:{name}",
                    )
                    continue
                if kind == "symlink":
                    _require(
                        member.issym() and member.linkname == link_target and member.size == 0,
                        f"archive_symlink_invalid:{name}",
                    )
                    continue
                _require(
                    member.isfile() and not member.issym() and member.linkname == "", f"archive_file_invalid:{name}"
                )
                stream = archive.extractfile(member)
                if stream is None:
                    raise LinuxArchiveError(f"archive_file_unreadable:{name}")
                observed = hashlib.sha256()
                size = 0
                with stream:
                    for chunk in iter(lambda source=stream: source.read(1024 * 1024), b""):
                        observed.update(chunk)
                        size += len(chunk)
                _require(size == member.size, f"archive_file_size_drift:{name}")
                _require(observed.hexdigest() == digest, f"archive_file_hash_drift:{name}")
    except (OSError, tarfile.TarError) as exc:
        raise LinuxArchiveError("archive_verification_failed") from exc


def _publish_no_replace(temporary: Path, destination: Path) -> None:
    try:
        os.link(temporary, destination)
    except FileExistsError as exc:
        raise LinuxArchiveError("archive_destination_exists") from exc
    except OSError as exc:
        raise LinuxArchiveError("archive_publish_failed") from exc
    try:
        temporary.unlink()
        if os.name == "posix":
            descriptor = os.open(destination.parent, os.O_RDONLY)
            try:
                os.fsync(descriptor)
            finally:
                os.close(descriptor)
    except OSError as exc:
        with contextlib.suppress(OSError):
            destination.unlink()
        raise LinuxArchiveError("archive_publish_durability_failed") from exc


def build_archive(
    *,
    contract_path: Path,
    artifact_id: str,
    payload_root: Path,
    destination: Path,
) -> dict[str, object]:
    manifest = read_contract(contract_path)
    artifact = select_artifact(manifest, artifact_id)
    payload_root = Path(os.path.abspath(payload_root))
    destination = Path(os.path.abspath(destination))
    _require(payload_root.name == artifact["sourceDirectory"], "payload_source_directory_name_mismatch")
    _require(destination.name == artifact["archiveName"], "archive_destination_name_mismatch")
    _require(
        destination != payload_root and payload_root not in destination.parents, "archive_destination_inside_payload"
    )
    _plain_directory_status(destination.parent, "archive_destination_parent")
    _require(not os.path.lexists(destination), "archive_destination_exists")
    temporary = destination.with_name(f".{destination.name}.tmp")
    _require(not os.path.lexists(temporary), "archive_temporary_exists")

    tree = _capture_tree(payload_root, artifact, manifest["generatedMetadata"])
    embedded_payload = _payload_manifest(manifest, artifact, tree)
    embedded_bytes = canonical_bytes(embedded_payload)
    generated = manifest["generatedMetadata"]
    checksums = _checksums(tree, str(generated["payloadManifestPath"]), embedded_bytes)
    temporary_created = False
    published = False
    try:
        _write_archive(payload_root, temporary, manifest, artifact, tree, embedded_bytes, checksums)
        temporary_created = True
        _verify_gzip_header(temporary, manifest["archive"])
        _verify_archive(temporary, manifest, artifact, tree, embedded_bytes, checksums)
        _require(_capture_tree(payload_root, artifact, generated) == tree, "payload_changed_during_archive")
        _publish_no_replace(temporary, destination)
        published = True
    except Exception:
        if temporary_created and os.path.lexists(temporary):
            with contextlib.suppress(OSError):
                temporary.unlink()
        if published:
            with contextlib.suppress(OSError):
                destination.unlink()
        raise
    result = {
        "schemaVersion": 1,
        "artifactId": artifact_id,
        "archiveName": destination.name,
        "bytes": destination.stat().st_size,
        "sha256": sha256_file(destination),
        "payloadManifestSha256": hashlib.sha256(embedded_bytes).hexdigest(),
        "checksumsSha256": hashlib.sha256(checksums).hexdigest(),
        "files": len(tree.files) + 2,
        "symlinks": len(tree.links),
        "directories": len(tree.directories) + 1,
    }
    return result


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--contract", type=Path, required=True)
    result.add_argument("--artifact", choices=tuple(_EXPECTED_ARTIFACT_SHAPES), required=True)
    result.add_argument("--payload-root", type=Path, required=True)
    result.add_argument("--output", type=Path, required=True)
    return result


def main(argv: Sequence[str] | None = None) -> int:
    try:
        args = parser().parse_args(argv)
        result = build_archive(
            contract_path=args.contract,
            artifact_id=args.artifact,
            payload_root=args.payload_root,
            destination=args.output,
        )
    except (LinuxArchiveError, OSError, KeyError, TypeError, ValueError, tarfile.TarError) as exc:
        print(f"linux archive failed closed: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
