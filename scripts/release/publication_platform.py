"""Retrieve only a selected platform, bound to the attested full candidate manifest."""

from __future__ import annotations

import re
import shutil
import stat
import subprocess
import zipfile
from pathlib import Path
from typing import Any

from scripts.release.publication_contract import (
    CHECKSUM_NAME,
    MANIFEST_NAME,
    canonical_digest,
    canonical_json,
    checksum_bytes,
    file_identity,
    package_names,
    read_object,
    require,
    verify_manifest,
    verify_origin,
)
from scripts.release.publication_http import GitHub, env_seconds
from scripts.release.publication_session import verify_candidate_jobs, verify_provenance


def _download(artifact: dict[str, Any], output: Path, names: set[str], *, repository: str, limit: int) -> None:
    archive = output / f"download-{artifact['id']}.zip"
    with archive.open("xb") as stream:
        result = subprocess.run(
            ["gh", "api", f"/repos/{repository}/actions/artifacts/{artifact['id']}/zip"],
            stdout=stream,
            stderr=subprocess.PIPE,
            timeout=env_seconds("DOCWEN_PUBLICATION_ARTIFACT_TIMEOUT", 600),
        )
    require(result.returncode == 0, "platform artifact download failed")
    require(f"sha256:{file_identity(archive)['sha256']}" == artifact["digest"], "downloaded artifact digest mismatch")
    with zipfile.ZipFile(archive) as bundle:
        entries = bundle.infolist()
        require(
            len(entries) == len(names) and {entry.filename for entry in entries} == names,
            "artifact ZIP inventory mismatch",
        )
        require(
            all(
                not entry.is_dir() and not stat.S_ISLNK(entry.external_attr >> 16) and entry.file_size <= limit
                for entry in entries
            ),
            "unexpected artifact member",
        )
        for entry in entries:
            with bundle.open(entry) as source, (output / entry.filename).open("xb") as target:
                shutil.copyfileobj(source, target)
    archive.unlink()


def _sibling(artifacts: list[dict[str, Any]], name: str, run_id: int) -> dict[str, Any]:
    matches = [item for item in artifacts if item.get("name") == name]
    require(len(matches) == 1, f"exact candidate artifact missing or ambiguous: {name}")
    item = matches[0]
    require(type(item.get("id")) is int and item["id"] > 0 and item.get("expired") is False, "artifact unavailable")
    require(item.get("workflow_run", {}).get("id") == run_id, "artifact belongs to another run")
    require(item.get("digest") == canonical_digest(str(item.get("digest", ""))), "artifact digest missing")
    return item


def fetch_platform(
    api: GitHub,
    *,
    output: Path,
    receipt: Path,
    artifact_id: int,
    digest: str,
    repository: str,
    version: str,
    commit: str,
    platform: str,
) -> dict[str, Any]:
    """No publication writes; a partial platform directory cannot pass the full release inventory."""
    require(platform in {"windows", "linux"}, "unknown candidate platform")
    names = set(package_names(version))
    selected = {name for name in names if name.endswith(".zip") == (platform == "windows")}
    prefix = f"/repos/{repository}"
    publication = api.get(f"{prefix}/actions/artifacts/{artifact_id}")
    require(publication.get("id") == artifact_id and publication.get("digest") == digest, "artifact identity mismatch")
    require(publication.get("expired") is False, "candidate artifact expired")
    match = re.fullmatch(r"docwen-publication-([1-9]\d*)-([1-9]\d*)", publication.get("name", ""))
    require(match is not None, "unexpected publication artifact name")
    assert match is not None
    run_id, attempt = map(int, match.groups())
    run = api.get(f"{prefix}/actions/runs/{run_id}/attempts/{attempt}")
    verify_candidate_jobs(api.get(f"{prefix}/actions/runs/{run_id}/attempts/{attempt}/jobs?per_page=100"))
    # Search only within the explicitly selected producer run, never earlier/latest runs.
    listing = api.get(f"{prefix}/actions/runs/{run_id}/artifacts?per_page=100")
    artifacts = listing.get("artifacts", [])
    require(listing.get("total_count") == len(artifacts), "incomplete candidate artifact response")
    metadata = _sibling(artifacts, f"docwen-publication-metadata-{run_id}-{attempt}", run_id)
    package = _sibling(artifacts, f"docwen-{platform}-{run_id}-{attempt}", run_id)
    require(
        type(metadata.get("size_in_bytes")) is int and 0 < metadata["size_in_bytes"] <= 2 * 1024**2,
        "metadata artifact too large",
    )
    require(output.parent.is_dir() and not output.is_symlink(), "candidate directory must be inside an owned directory")
    require(receipt.parent.is_dir() and not receipt.is_symlink(), "receipt must be inside an owned directory")
    require(not receipt.resolve().is_relative_to(output.resolve()), "receipt must be outside candidate inventory")
    reused = output.exists()
    if not reused:
        output.mkdir()
        _download(metadata, output, {MANIFEST_NAME, CHECKSUM_NAME}, repository=repository, limit=1024**2)
    require(
        output.is_dir() and not getattr(output.stat(), "st_file_attributes", 0) & 0x400,
        "reparse candidate directory rejected",
    )
    for name in (MANIFEST_NAME, CHECKSUM_NAME):
        require(file_identity(output / name)["bytes"] <= 1024**2, "candidate metadata too large")
    manifest = read_object(output / MANIFEST_NAME)
    verify_manifest(manifest, repository=repository, version=version, commit=commit)
    verify_origin(manifest, run, publication, digest=digest)
    require((output / CHECKSUM_NAME).read_bytes() == checksum_bytes(manifest["assets"]), "checksum inventory mismatch")
    verify_provenance(output, manifest, (MANIFEST_NAME, CHECKSUM_NAME))
    if not reused:
        _download(package, output, selected, repository=repository, limit=2 * 1024**3)
    require(
        {path.name for path in output.iterdir()} == selected | {MANIFEST_NAME, CHECKSUM_NAME},
        "unexpected platform candidate files",
    )
    for name in selected:
        require(file_identity(output / name) == manifest["assets"][name], f"candidate asset hash mismatch: {name}")
    verify_provenance(output, manifest, tuple(sorted(selected)))
    result = {
        "schema": "docwen-platform-candidate-v1",
        "stage": "platform-candidate-verified",
        "repository": repository,
        "version": version,
        "sourceCommit": commit,
        "platform": platform,
        "artifactId": artifact_id,
        "artifactDigest": digest,
        "origin": manifest["origin"],
        "manifestSha256": file_identity(output / MANIFEST_NAME)["sha256"],
        "metadataArtifact": {key: metadata[key] for key in ("id", "digest")},
        "platformArtifact": {key: package[key] for key in ("id", "digest")},
        "assets": {name: manifest["assets"][name] for name in sorted(selected)},
        "provenance": "verified",
        "scope": "Selected platform bytes only; no native acceptance, other-platform download or publication claim.",
    }
    content = canonical_json(result)
    if receipt.exists():
        require(receipt.read_bytes() == content, "platform receipt identity mismatch")
    else:
        with receipt.open("xb") as stream:
            stream.write(content)
    return {**result, "reused": reused}
