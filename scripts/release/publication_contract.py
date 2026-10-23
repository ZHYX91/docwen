"""Exact source and artifact inventory used from candidate through publication."""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import stat
from pathlib import Path
from typing import Any

SCHEMA = "docwen-publication-candidate-v2"
MANIFEST_NAME = "candidate.json"
CHECKSUM_NAME = "SHA256SUMS.txt"
WORKFLOW = ".github/workflows/release.yml"


class PublicationError(RuntimeError):
    """A failed source, package, or publication boundary."""


def command_diagnostic(subject: str, exit_code: int | str, stderr: str | bytes | None) -> str:
    """Keep a bounded failure reason without credentials or transport URLs."""
    message = stderr.decode("utf-8", errors="replace") if isinstance(stderr, bytes) else (stderr or "")
    # Redact before truncating: a secret may cross the final length boundary.
    for key in (
        "GH_TOKEN",
        "GITHUB_TOKEN",
        "GH_ENTERPRISE_TOKEN",
        "GITHUB_ENTERPRISE_TOKEN",
        "HTTPS_PROXY",
        "HTTP_PROXY",
        "ALL_PROXY",
    ):
        if secret := os.environ.get(key):
            message = message.replace(secret, "[redacted]")
    message = re.sub(r"(?i)\b[a-z][a-z0-9+.-]*://[^\s<>\"']+", "[redacted-url]", message)
    message = re.sub(r"(?i)\b(?:bearer|basic)\s+\S+", "[redacted-authorization]", message)
    message = re.sub(r"\b(?:gh[pousr]_[A-Za-z0-9_]+|github_pat_[A-Za-z0-9_]+)", "[redacted-token]", message)
    message = re.sub(
        r"""(?i)(["']?)\b((?:access[_-]?)?token|authorization|password|client[_-]?secret|signature|sig)\1\s*[:=]\s*(?:"[^"]*"|'[^']*'|[^\s,;]+)""",
        r"\2=[redacted]",
        message,
    )
    # Escape control characters through JSON, including terminal escapes/newlines.
    message = "".join(character if character.isprintable() else " " for character in message)
    if len(message) > 1024:
        message = message[:1024] + " [truncated]"
    return f"{subject}; exit_code={exit_code}; stderr={json.dumps(message, ensure_ascii=True)}"


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


def package_names(version: str) -> tuple[str, ...]:
    require(bool(re.fullmatch(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)", version)), "invalid release version")
    return (
        "DocWen-windows-x64.zip",
        f"DocWen-{version}-linux-x64.tar.gz",
        f"DocWenCLI-{version}-linux-x64.tar.gz",
        f"DocWen-{version}-macos-arm64.tar.gz",
        f"DocWenCLI-{version}-macos-arm64.tar.gz",
    )


def package_platform(name: str, version: str) -> str:
    """Resolve an exact declared asset, never its archive extension."""
    names = package_names(version)
    require(name in names, "unknown package asset")
    return dict(zip(names, ("windows", "linux", "linux", "macos", "macos"), strict=True))[name]


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
    builds: Path,
    output: Path,
    *,
    repo: str,
    version: str,
    commit: str,
    run_id: int,
    attempt: int,
    source: Path,
    source_ref: str,
) -> dict[str, Any]:
    require(bool(re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repo)), "invalid repository")
    require(bool(re.fullmatch(r"[0-9a-f]{40}", commit)), "invalid source commit")
    require(run_id > 0 and attempt > 0, "invalid candidate run identity")
    require(not output.exists(), "publication output must be new")
    records: dict[str, dict[str, Any]] = {}
    sources: dict[str, Path] = {}
    for name in package_names(version):
        platform = package_platform(name, version)
        package = builds / platform / name
        records[name] = file_identity(package)
        sources[name] = package
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
                "release/macos-production-manifest.v1.json",
            )
        },
        "origin": {"workflow": WORKFLOW, "runId": run_id, "runAttempt": attempt, "sourceRef": source_ref},
        "assets": records,
    }
    # Check every input before making the candidate visible. Publication reuses
    # these exact packages; reproducing builds is an optional engineering check.
    output.mkdir(parents=True)
    for name, path in sources.items():
        shutil.copyfile(path, output / name)
    (output / CHECKSUM_NAME).write_bytes(checksum_bytes(records))
    (output / MANIFEST_NAME).write_bytes(canonical_json(manifest))
    verify_inventory(output, repository=repo, version=version, commit=commit)
    return manifest


def verify_manifest(manifest: dict[str, Any], *, repository: str, version: str, commit: str) -> None:
    """Validate the full candidate identity without requiring every platform locally."""
    require(manifest.get("schema") == SCHEMA, "unsupported candidate manifest")
    require(
        (manifest.get("repository"), manifest.get("version"), manifest.get("sourceCommit"))
        == (repository, version, commit),
        "candidate source identity mismatch",
    )
    names = set(package_names(version))
    require(set(manifest.get("assets", {})) == names, "candidate asset inventory mismatch")
    for identity in manifest["assets"].values():
        require(isinstance(identity, dict), "invalid candidate asset identity")
        require(type(identity.get("bytes")) is int and 0 < identity["bytes"] <= 2 * 1024**3, "invalid asset size")
        require(bool(re.fullmatch(r"[0-9a-f]{64}", str(identity.get("sha256", "")))), "invalid asset digest")
    origin = manifest.get("origin", {})
    require(origin.get("workflow") == WORKFLOW, "unexpected builder workflow")
    require(type(origin.get("runId")) is int and origin["runId"] > 0, "candidate run missing")
    require(type(origin.get("runAttempt")) is int and origin["runAttempt"] > 0, "candidate attempt missing")
    source_ref = origin.get("sourceRef")
    require(
        isinstance(source_ref, str)
        and (
            source_ref == f"refs/tags/{version}"
            or (source_ref.startswith("refs/heads/") and len(source_ref) > len("refs/heads/"))
        ),
        "candidate source ref missing or invalid",
    )
    expected_inputs = {
        "uv.lock",
        "release/windows-production-manifest.v1.json",
        "release/linux-production-manifest.v1.json",
        "release/macos-production-manifest.v1.json",
    }
    require(set(manifest.get("sourceInputs", {})) == expected_inputs, "source input identities missing")


def verify_inventory(directory: Path, *, repository: str, version: str, commit: str) -> dict[str, Any]:
    manifest = read_object(directory / MANIFEST_NAME)
    verify_manifest(manifest, repository=repository, version=version, commit=commit)
    names = set(package_names(version))
    require(
        {path.name for path in directory.iterdir()} == names | {CHECKSUM_NAME, MANIFEST_NAME},
        "unexpected candidate files",
    )
    for name in names:
        require(file_identity(directory / name) == manifest["assets"][name], f"candidate asset hash mismatch: {name}")
    require(
        (directory / CHECKSUM_NAME).read_bytes() == checksum_bytes(manifest["assets"]), "checksum inventory mismatch"
    )
    return manifest


def publication_assets(directory: Path, manifest: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {**manifest["assets"], CHECKSUM_NAME: file_identity(directory / CHECKSUM_NAME)}


def verify_origin(manifest: dict[str, Any], run: dict[str, Any], artifact: dict[str, Any], *, digest: str) -> None:
    origin = manifest["origin"]
    require(run.get("id") == origin["runId"] and run.get("run_attempt") == origin["runAttempt"], "wrong candidate run")
    # The candidate is consumed later in its own run. A completed run can also
    # have failed only during publication; exact producer jobs are checked separately.
    require(run.get("status") in {"in_progress", "completed"}, "candidate run has not started")
    require(
        run.get("head_sha") == manifest["sourceCommit"] and run.get("path") == WORKFLOW, "candidate source mismatch"
    )
    require(run.get("event") in {"push", "workflow_dispatch"}, "unexpected candidate trigger")
    source_ref = origin["sourceRef"]
    require(run.get("head_branch") == source_ref.split("/", 2)[2], "candidate source ref mismatch")
    require(
        run.get("event") != "push" or source_ref == f"refs/tags/{manifest['version']}",
        "push candidate requires its numeric tag",
    )
    require(run.get("repository", {}).get("full_name") == manifest["repository"], "candidate repository mismatch")
    require(artifact.get("expired") is False, "candidate artifact expired")
    require(artifact.get("workflow_run", {}).get("id") == origin["runId"], "artifact belongs to another run")
    require(
        artifact.get("name") == f"docwen-publication-{origin['runId']}-{origin['runAttempt']}", "artifact name mismatch"
    )
    require(
        bool(re.fullmatch(r"sha256:[0-9a-f]{64}", digest)) and artifact.get("digest") == digest,
        "artifact digest mismatch",
    )
