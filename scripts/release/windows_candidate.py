"""Attested Windows producer handoff, independent of the full release inventory."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.release.publication_contract import (
    WORKFLOW,
    canonical_json,
    file_identity,
    package_names,
    read_object,
    require,
)
from scripts.release.publication_download import atomic_record, pending_record, regular_file, require_record_filesystem
from scripts.release.publication_http import GitHub
from scripts.release.publication_session import candidate_jobs, verify_provenance

NAME = "windows-candidate.json"
ARCHIVE = "DocWen-windows-x64.zip"
SCHEMA = "docwen-windows-producer-v1"
SOURCE_JOBS = {
    "Source identity",
    "source-checks",
    "Tests on windows-latest",
    "Tests on ubuntu-24.04",
    "Tests on macos-15",
    "Required checks",
    "Build Windows packages",
}


def validate(record: dict[str, Any], *, repository: str, version: str, commit: str) -> None:
    package_names(version)
    require(record.get("schema") == SCHEMA, "unsupported Windows producer manifest")
    require(
        (record.get("repository"), record.get("version"), record.get("sourceCommit")) == (repository, version, commit),
        "Windows producer source mismatch",
    )
    require(bool(re.fullmatch(r"[0-9a-f]{40}", commit)), "invalid source commit")
    require(set(record.get("assets", {})) == {ARCHIVE}, "Windows producer inventory mismatch")
    asset = record["assets"][ARCHIVE]
    require(
        isinstance(asset, dict)
        and type(asset.get("bytes")) is int
        and 0 < asset["bytes"] <= 2 * 1024**3
        and bool(re.fullmatch(r"[0-9a-f]{64}", str(asset.get("sha256", "")))),
        "invalid Windows asset",
    )
    origin = record.get("origin", {})
    require(origin.get("workflow") == WORKFLOW, "unexpected Windows producer workflow")
    for field in ("runId", "runAttempt"):
        require(type(origin.get(field)) is int and origin[field] > 0, "invalid Windows producer run")
    ref = origin.get("sourceRef", "")
    require(
        isinstance(ref, str) and (ref == f"refs/tags/{version}" or (ref.startswith("refs/heads/") and len(ref) > 11)),
        "invalid Windows producer ref",
    )


def create(
    directory: Path, *, repository: str, version: str, commit: str, run_id: int, attempt: int, source_ref: str
) -> dict[str, Any]:
    record = {
        "schema": SCHEMA,
        "repository": repository,
        "version": version,
        "sourceCommit": commit,
        "origin": {"workflow": WORKFLOW, "runId": run_id, "runAttempt": attempt, "sourceRef": source_ref},
        "assets": {ARCHIVE: file_identity(directory / ARCHIVE)},
    }
    validate(record, repository=repository, version=version, commit=commit)
    with (directory / NAME).open("xb") as stream:
        stream.write(canonical_json(record))
    return record


def fetch(
    api: GitHub,
    *,
    output: Path,
    receipt: Path,
    artifact_id: int,
    digest: str,
    repository: str,
    version: str,
    commit: str,
) -> dict[str, Any]:
    # Import here to keep the existing platform downloader reusable without a cycle.
    from scripts.release.publication_platform import _download, _sibling

    prefix = f"/repos/{repository}"
    artifact = api.get(f"{prefix}/actions/artifacts/{artifact_id}")
    require(
        artifact.get("id") == artifact_id and artifact.get("digest") == digest and artifact.get("expired") is False,
        "Windows artifact identity mismatch",
    )
    match = re.fullmatch(r"docwen-windows-([1-9]\d*)-([1-9]\d*)", artifact.get("name", ""))
    require(match is not None, "exact Windows producer artifact required")
    assert match is not None
    run_id, attempt = map(int, match.groups())
    require(artifact.get("workflow_run", {}).get("id") == run_id, "Windows artifact run mismatch")
    run = api.get(f"{prefix}/actions/runs/{run_id}/attempts/{attempt}")
    require(
        run.get("id") == run_id
        and run.get("run_attempt") == attempt
        and run.get("head_sha") == commit
        and run.get("path") == WORKFLOW
        and run.get("event") in {"push", "workflow_dispatch"}
        and run.get("repository", {}).get("full_name") == repository
        and run.get("status") in {"in_progress", "completed"},
        "Windows producer run mismatch",
    )
    jobs = candidate_jobs(api, run_id=run_id, attempt=attempt, commit=commit)
    for name in SOURCE_JOBS:
        matches = [job for job in jobs["jobs"] if job.get("name", "").split(" / ")[-1] == name]
        require(
            len(matches) == 1 and matches[0].get("conclusion") == "success",
            f"required Windows producer check did not pass: {name}",
        )
    producer = next(job for job in jobs["jobs"] if job["name"] == "Build Windows packages")
    require(producer.get("run_attempt", attempt) == attempt, "Windows producer attempt mismatch")
    listing = api.get(f"{prefix}/actions/runs/{run_id}/artifacts?per_page=100")
    require(listing.get("total_count") == len(listing.get("artifacts", [])), "incomplete artifact listing")
    metadata = _sibling(listing["artifacts"], f"docwen-windows-metadata-{run_id}-{attempt}", run_id)
    require(
        type(metadata.get("size_in_bytes")) is int and 0 < metadata["size_in_bytes"] <= 1024**2,
        "Windows metadata too large",
    )
    require(
        output.parent.is_dir()
        and receipt.parent.is_dir()
        and not output.is_symlink()
        and not receipt.is_symlink()
        and not receipt.resolve().is_relative_to(output.resolve()),
        "Windows candidate requires owned output and external receipt",
    )
    require_record_filesystem(receipt.parent)
    if not output.exists():
        output.mkdir()
    require(
        output.is_dir() and not getattr(output.stat(), "st_file_attributes", 0) & 0x400,
        "reparse Windows candidate directory rejected",
    )
    marker = output / "transfer.json"
    binding = {
        "artifactId": artifact_id,
        "digest": digest,
        "repository": repository,
        "version": version,
        "commit": commit,
        "metadata": metadata["id"],
    }
    if marker.exists():
        regular_file(marker)
        require(read_object(marker) == binding, "Windows transfer identity mismatch")
    elif receipt.exists():
        regular_file(receipt)
        prior = read_object(receipt)
        require(
            prior.get("artifactId") == artifact_id and prior.get("artifactDigest") == digest,
            "Windows receipt identity mismatch",
        )
    else:
        pending = pending_record(marker, canonical_json(binding))
        require({path.name for path in output.iterdir()} <= {pending.name}, "unowned Windows candidate directory")
    atomic_record(marker, canonical_json(binding))
    allowed = {ARCHIVE, NAME, marker.name, ARCHIVE + ".extracting", NAME + ".extracting"}
    for item in (artifact, metadata):
        state = output / f"download-{item['id']}.json"
        identity = {
            "repository": repository,
            "id": item["id"],
            "digest": item["digest"],
            "bytes": item["size_in_bytes"],
        }
        allowed |= {f"download-{item['id']}.zip", state.name, pending_record(state, canonical_json(identity)).name}
    require({path.name for path in output.iterdir()} <= allowed, "unexpected Windows transfer files")
    if not (output / NAME).is_file() or (output / f"download-{metadata['id']}.json").exists():
        _download(api, metadata, output, {NAME}, limit=1024**2)
    require(file_identity(output / NAME)["bytes"] <= 1024**2, "Windows metadata too large")
    record = read_object(output / NAME)
    validate(record, repository=repository, version=version, commit=commit)
    origin = record["origin"]
    require(
        (origin["runId"], origin["runAttempt"]) == (run_id, attempt)
        and origin["sourceRef"].split("/", 2)[2] == run.get("head_branch"),
        "Windows manifest origin mismatch",
    )
    verify_provenance(output, record, (NAME,))
    if not (output / ARCHIVE).is_file() or (output / f"download-{artifact['id']}.json").exists():
        _download(api, artifact, output, {ARCHIVE}, limit=2 * 1024**3)
    require(file_identity(output / ARCHIVE) == record["assets"][ARCHIVE], "Windows archive hash mismatch")
    verify_provenance(output, record, (ARCHIVE,))
    require({path.name for path in output.iterdir()} == {NAME, ARCHIVE, marker.name}, "unexpected Windows inventory")
    result = {
        "schema": "docwen-platform-candidate-v1",
        "stage": "platform-candidate-verified",
        "candidateKind": "windows-producer",
        "repository": repository,
        "version": version,
        "sourceCommit": commit,
        "platform": "windows",
        "artifactId": artifact_id,
        "artifactDigest": digest,
        "origin": origin,
        "manifestSha256": file_identity(output / NAME)["sha256"],
        "metadataArtifact": {key: metadata[key] for key in ("id", "digest", "size_in_bytes")},
        "platformArtifact": {key: artifact[key] for key in ("id", "digest", "size_in_bytes")},
        "assets": record["assets"],
        "provenance": "verified",
        "scope": "Windows producer only; no full-release, Linux/macOS or native acceptance claim.",
    }
    atomic_record(receipt, canonical_json(result))
    marker.unlink()
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    for field in ("repository", "version", "commit", "source-ref"):
        parser.add_argument("--" + field, required=True)
    for field in ("run-id", "attempt"):
        parser.add_argument("--" + field, type=int, required=True)
    parser.add_argument("--directory", type=Path, required=True)
    args = parser.parse_args()
    create(
        args.directory,
        repository=args.repository,
        version=args.version,
        commit=args.commit,
        run_id=args.run_id,
        attempt=args.attempt,
        source_ref=args.source_ref,
    )


if __name__ == "__main__":
    main()
