"""Retrieve only a selected platform, bound to the attested full candidate manifest."""

from __future__ import annotations

import hashlib
import os
import re
import shutil
import stat
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
from scripts.release.publication_download import download_artifact, finish_download, regular_file
from scripts.release.publication_http import GitHub
from scripts.release.publication_session import verify_candidate_jobs, verify_provenance


def _download(api: GitHub, artifact: dict[str, Any], output: Path, names: set[str], *, limit: int) -> None:
    archive = output / f"download-{artifact['id']}.zip"
    download_artifact(api, artifact, archive)
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
            destination = output / entry.filename
            if destination.exists():
                regular_file(destination)
                with bundle.open(entry) as source:
                    require(
                        file_identity(destination)
                        == {"bytes": entry.file_size, "sha256": hashlib.file_digest(source, "sha256").hexdigest()},
                        "previous extraction differs from verified archive",
                    )
                temporary = output / (entry.filename + ".extracting")
                if temporary.exists():
                    regular_file(temporary)
                    temporary.unlink()
                continue
            temporary = output / (entry.filename + ".extracting")
            if temporary.exists():
                regular_file(temporary)
                temporary.unlink()
            # The verified transfer sidecar owns this scratch name. Never expose
            # a partially written member under its final candidate filename.
            with bundle.open(entry) as source, temporary.open("xb") as target:
                shutil.copyfileobj(source, target)
            require(temporary.stat().st_size == entry.file_size, "incomplete extracted member")
            os.link(temporary, destination)
            temporary.unlink()
    finish_download(archive)


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
    marker = output / "transfer.json"
    transfer = {
        "artifactId": artifact_id,
        "digest": digest,
        "repository": repository,
        "version": version,
        "commit": commit,
        "platform": platform,
        "siblings": {
            "metadata": {key: metadata[key] for key in ("id", "digest", "size_in_bytes")},
            "package": {key: package[key] for key in ("id", "digest", "size_in_bytes")},
        },
    }
    if not reused:
        output.mkdir()
        with marker.open("xb") as stream:
            stream.write(canonical_json(transfer))
    require(
        output.is_dir() and not getattr(output.stat(), "st_file_attributes", 0) & 0x400,
        "reparse candidate directory rejected",
    )
    recovering = marker.exists()
    require(recovering or receipt.is_file(), "candidate cache has no identity receipt or transfer journal")
    if recovering:
        regular_file(marker)
        require(read_object(marker) == transfer, "candidate transfer identity mismatch")
        allowed = selected | {MANIFEST_NAME, CHECKSUM_NAME, marker.name}
        allowed |= {name + ".extracting" for name in selected | {MANIFEST_NAME, CHECKSUM_NAME}}
        for item in (metadata, package):
            allowed |= {f"download-{item['id']}.zip", f"download-{item['id']}.json"}
        require({path.name for path in output.iterdir()} <= allowed, "unexpected candidate recovery files")
        if (output / f"download-{metadata['id']}.json").exists() or not all(
            (output / name).is_file() for name in (MANIFEST_NAME, CHECKSUM_NAME)
        ):
            _download(api, metadata, output, {MANIFEST_NAME, CHECKSUM_NAME}, limit=1024**2)
    for name in (MANIFEST_NAME, CHECKSUM_NAME):
        require(file_identity(output / name)["bytes"] <= 1024**2, "candidate metadata too large")
    manifest = read_object(output / MANIFEST_NAME)
    verify_manifest(manifest, repository=repository, version=version, commit=commit)
    verify_origin(manifest, run, publication, digest=digest)
    require((output / CHECKSUM_NAME).read_bytes() == checksum_bytes(manifest["assets"]), "checksum inventory mismatch")
    verify_provenance(output, manifest, (MANIFEST_NAME, CHECKSUM_NAME))
    if recovering and (
        (output / f"download-{package['id']}.json").exists() or not all((output / name).is_file() for name in selected)
    ):
        _download(api, package, output, selected, limit=2 * 1024**3)
    require(
        {path.name for path in output.iterdir()}
        == selected | {MANIFEST_NAME, CHECKSUM_NAME} | ({marker.name} if recovering else set()),
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
        "metadataArtifact": {key: metadata[key] for key in ("id", "digest", "size_in_bytes")},
        "platformArtifact": {key: package[key] for key in ("id", "digest", "size_in_bytes")},
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
    if recovering:
        marker.unlink()
    return {**result, "reused": reused}
