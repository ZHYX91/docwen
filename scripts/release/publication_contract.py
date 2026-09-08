"""Exact source and artifact inventory used from preflight through publication."""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import stat
from pathlib import Path
from typing import Any

SCHEMA = "docwen-publication-candidate-v1"
MANIFEST_NAME = "candidate.json"
CHECKSUM_NAME = "SHA256SUMS.txt"
WORKFLOW = ".github/workflows/release.yml"


class PublicationError(RuntimeError):
    """A failed source, package, or publication boundary."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise PublicationError(message)


def file_identity(path: Path) -> dict[str, Any]:
    before = path.lstat()
    require(stat.S_ISREG(before.st_mode) and not path.is_symlink(), f"regular file required: {path.name}")
    require(not getattr(before, "st_file_attributes", 0) & 0x400, f"reparse file rejected: {path.name}")
    with path.open("rb") as stream:
        digest = hashlib.file_digest(stream, "sha256").hexdigest()
    after = path.lstat()
    require(
        (before.st_dev, before.st_ino, before.st_size, before.st_mtime_ns)
        == (after.st_dev, after.st_ino, after.st_size, after.st_mtime_ns),
        f"file changed while hashing: {path.name}",
    )
    return {"bytes": after.st_size, "sha256": digest}


def package_names(version: str) -> tuple[str, str, str]:
    require(bool(re.fullmatch(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)", version)), "invalid release version")
    return (
        "DocWen-windows-x64.zip",
        f"DocWen-{version}-linux-x64.tar.gz",
        f"DocWenCLI-{version}-linux-x64.tar.gz",
    )


def canonical_json(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2) + "\n").encode("utf-8")


def canonical_digest(value: str) -> str:
    digest = value.removeprefix("sha256:")
    require(bool(re.fullmatch(r"[0-9a-f]{64}", digest)), "invalid artifact SHA-256 digest")
    return f"sha256:{digest}"


def read_object(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    require(isinstance(value, dict), f"JSON object required: {path.name}")
    return value


def checksum_bytes(assets: dict[str, dict[str, Any]]) -> bytes:
    return "".join(f"{assets[name]['sha256']}  {name}\n" for name in sorted(assets)).encode("ascii")


def assemble(
    builds: Path, output: Path, *, repo: str, version: str, commit: str, run_id: int, attempt: int, source: Path
) -> dict[str, Any]:
    require(bool(re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repo)), "invalid repository")
    require(bool(re.fullmatch(r"[0-9a-f]{40}", commit)), "invalid source commit")
    require(run_id > 0 and attempt > 0, "invalid preflight run identity")
    require(not output.exists(), "publication output must be new")
    records: dict[str, dict[str, Any]] = {}
    sources: dict[str, Path] = {}
    for name in package_names(version):
        platform = "windows" if name.endswith(".zip") else "linux"
        first = builds / f"{platform}-a" / name
        require(
            file_identity(first) == file_identity(builds / f"{platform}-b" / name), f"independent builds differ: {name}"
        )
        records[name] = file_identity(first)
        sources[name] = first
    msix_records = []
    for replica in ("a", "b"):
        root = builds / f"windows-{replica}"
        record = read_object(root / "DocWen-windows-x64.msix.json")
        observed = file_identity(root / "DocWen-windows-x64.msix")
        require(all(record.get(key) == value for key, value in observed.items()), "MSIX bytes differ from its receipt")
        require(record.get("sourceVersion") == version, "MSIX source version mismatch")
        msix_records.append(record)
    require(
        {key: value for key, value in msix_records[0].items() if key != "sha256"}
        == {key: value for key, value in msix_records[1].items() if key != "sha256"},
        "independent MSIX content identities differ",
    )
    manifest = {
        "schema": SCHEMA,
        "repository": repo,
        "version": version,
        "sourceCommit": commit,
        "sourceInputs": {
            name: file_identity(source / name)
            for name in (
                "uv.lock",
                "release/windows-production-manifest.v1.json",
                "release/linux-production-manifest.v1.json",
            )
        },
        "origin": {"workflow": WORKFLOW, "runId": run_id, "runAttempt": attempt},
        "assets": records,
        "store": msix_records[0],
    }
    # Compare every input before making the candidate visible.
    output.mkdir(parents=True)
    for name, path in sources.items():
        shutil.copyfile(path, output / name)
    (output / CHECKSUM_NAME).write_bytes(checksum_bytes(records))
    (output / MANIFEST_NAME).write_bytes(canonical_json(manifest))
    verify_inventory(output, repository=repo, version=version, commit=commit)
    return manifest


def verify_inventory(directory: Path, *, repository: str, version: str, commit: str) -> dict[str, Any]:
    manifest = read_object(directory / MANIFEST_NAME)
    require(manifest.get("schema") == SCHEMA, "unsupported candidate manifest")
    require(
        (manifest.get("repository"), manifest.get("version"), manifest.get("sourceCommit"))
        == (repository, version, commit),
        "candidate source identity mismatch",
    )
    names = set(package_names(version))
    require(set(manifest.get("assets", {})) == names, "candidate asset inventory mismatch")
    require(
        {path.name for path in directory.iterdir()} == names | {CHECKSUM_NAME, MANIFEST_NAME},
        "unexpected candidate files",
    )
    for name in names:
        require(file_identity(directory / name) == manifest["assets"][name], f"candidate asset hash mismatch: {name}")
    require(
        (directory / CHECKSUM_NAME).read_bytes() == checksum_bytes(manifest["assets"]), "checksum inventory mismatch"
    )
    origin = manifest.get("origin", {})
    require(origin.get("workflow") == WORKFLOW, "unexpected builder workflow")
    require(type(origin.get("runId")) is int and origin["runId"] > 0, "preflight run missing")
    require(type(origin.get("runAttempt")) is int and origin["runAttempt"] > 0, "preflight attempt missing")
    expected_inputs = {
        "uv.lock",
        "release/windows-production-manifest.v1.json",
        "release/linux-production-manifest.v1.json",
    }
    require(set(manifest.get("sourceInputs", {})) == expected_inputs, "source input identities missing")
    return manifest


def publication_assets(directory: Path, manifest: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {**manifest["assets"], CHECKSUM_NAME: file_identity(directory / CHECKSUM_NAME)}


def verify_origin(manifest: dict[str, Any], run: dict[str, Any], artifact: dict[str, Any], *, digest: str) -> None:
    origin = manifest["origin"]
    require(run.get("id") == origin["runId"] and run.get("run_attempt") == origin["runAttempt"], "wrong preflight run")
    require(run.get("status") == "completed" and run.get("conclusion") == "success", "preflight did not pass")
    require(
        run.get("head_sha") == manifest["sourceCommit"] and run.get("path") == WORKFLOW, "preflight source mismatch"
    )
    require(run.get("event") == "workflow_dispatch", "preflight must be an explicit build run")
    require(run.get("repository", {}).get("full_name") == manifest["repository"], "preflight repository mismatch")
    require(artifact.get("expired") is False, "preflight artifact expired")
    require(artifact.get("workflow_run", {}).get("id") == origin["runId"], "artifact belongs to another run")
    require(
        artifact.get("name") == f"docwen-publication-{origin['runId']}-{origin['runAttempt']}", "artifact name mismatch"
    )
    require(
        bool(re.fullmatch(r"sha256:[0-9a-f]{64}", digest)) and artifact.get("digest") == digest,
        "artifact digest mismatch",
    )
