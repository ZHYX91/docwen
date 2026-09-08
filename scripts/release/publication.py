"""CLI for assembling, retrieving and publishing an already verified candidate."""

from __future__ import annotations

import argparse
import io
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.release.publication_contract import (
    CHECKSUM_NAME,
    MANIFEST_NAME,
    PublicationError,
    assemble,
    canonical_digest,
    canonical_json,
    file_identity,
    package_names,
    require,
    verify_inventory,
    verify_origin,
)
from scripts.release.publication_http import GitHub
from scripts.release.publication_session import ReleaseSession, verify_preflight_jobs


def restore_progress(api: GitHub, *, artifact_id: int, digest: str, receipt: Path, commit: str) -> None:
    prefix = f"/repos/{api.repository}"
    metadata = api.get(f"{prefix}/actions/artifacts/{artifact_id}")
    require(
        metadata.get("id") == artifact_id and metadata.get("expired") is False and metadata.get("digest") == digest,
        "progress artifact identity mismatch",
    )
    match = re.fullmatch(r"docwen-publication-progress-([1-9]\d*)-([1-9]\d*)", metadata.get("name", ""))
    require(match is not None, "unexpected progress artifact name")
    assert match is not None
    run_id, attempt = map(int, match.groups())
    require(metadata.get("workflow_run", {}).get("id") == run_id, "progress artifact belongs to another run")
    run = api.get(f"{prefix}/actions/runs/{run_id}/attempts/{attempt}")
    require(
        run.get("status") == "completed"
        and run.get("head_sha") == commit
        and run.get("path") == ".github/workflows/release.yml",
        "progress owner is active or has different source",
    )
    response = subprocess.run(
        ["gh", "api", f"{prefix}/actions/artifacts/{artifact_id}/zip"], capture_output=True, timeout=60
    )
    require(response.returncode == 0, "progress artifact download failed")
    import hashlib

    require(f"sha256:{hashlib.sha256(response.stdout).hexdigest()}" == digest, "progress archive digest mismatch")
    with zipfile.ZipFile(io.BytesIO(response.stdout)) as archive:
        require(archive.namelist() == ["publication-progress.json"], "progress ZIP inventory mismatch")
        require(archive.getinfo("publication-progress.json").file_size <= 65536, "progress receipt exceeds limit")
        state = json.loads(archive.read("publication-progress.json"))
    require(isinstance(state, dict) and state.get("sourceCommit") == commit, "progress receipt source mismatch")
    content = canonical_json(state)
    if receipt.exists():
        require(receipt.read_bytes() == content, "existing progress receipt differs")
    else:
        with receipt.open("xb") as stream:
            stream.write(content)


def release_notes(path: Path, version: str) -> str:
    text = path.read_text(encoding="utf-8")
    sections = re.split(r"(?m)^## ", text)
    matches = [section for section in sections[1:] if re.match(rf"{re.escape(version)}(?:\s|$)", section)]
    require(len(matches) == 1, "exact version release notes missing or duplicated")
    return "## " + matches[0].strip()


def fetch_candidate(
    api: GitHub,
    *,
    output: Path,
    artifact_id: int,
    digest: str,
    repository: str,
    version: str,
    commit: str,
) -> None:
    require(not output.exists() and output.parent.is_dir(), "candidate output must be new inside an owned directory")
    metadata = api.get(f"/repos/{repository}/actions/artifacts/{artifact_id}")
    require(
        metadata.get("id") == artifact_id and metadata.get("digest") == digest and metadata.get("expired") is False,
        "artifact identity mismatch",
    )
    output.mkdir()
    archive = output / "download.zip"
    with archive.open("xb") as stream:
        result = subprocess.run(
            ["gh", "api", f"/repos/{repository}/actions/artifacts/{artifact_id}/zip"],
            stdout=stream,
            stderr=subprocess.PIPE,
            timeout=600,
        )
    require(result.returncode == 0, "candidate artifact download failed")
    require(f"sha256:{file_identity(archive)['sha256']}" == digest, "downloaded artifact digest mismatch")
    expected = set(package_names(version)) | {MANIFEST_NAME, CHECKSUM_NAME}
    with zipfile.ZipFile(archive) as bundle:
        entries = bundle.infolist()
        require(
            len(entries) == len(expected) and {entry.filename for entry in entries} == expected,
            "artifact ZIP inventory mismatch",
        )
        require(
            all(entry.file_size <= 2 * 1024**3 and not entry.is_dir() for entry in entries),
            "unexpected artifact member",
        )
        for entry in entries:
            with bundle.open(entry) as source, (output / entry.filename).open("xb") as target:
                shutil.copyfileobj(source, target)
    archive.unlink()
    manifest = verify_inventory(output, repository=repository, version=version, commit=commit)
    origin = manifest["origin"]
    run = api.get(f"/repos/{repository}/actions/runs/{origin['runId']}")
    verify_origin(manifest, run, metadata, digest=digest)
    verify_preflight_jobs(
        api.get(f"/repos/{repository}/actions/runs/{origin['runId']}/attempts/{origin['runAttempt']}/jobs?per_page=100")
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="operation", required=True)
    for operation in ("assemble", "fetch", "publish", "verify"):
        command = subparsers.add_parser(operation)
        command.add_argument("--repository", required=True)
        command.add_argument("--version", required=True)
        command.add_argument("--commit", required=True)
        command.add_argument("--directory", type=Path, required=True)
        if operation == "assemble":
            command.add_argument("--builds", type=Path, required=True)
            command.add_argument("--source", type=Path, required=True)
            command.add_argument("--run-id", type=int, required=True)
            command.add_argument("--attempt", type=int, required=True)
        else:
            command.add_argument("--artifact-id", type=int, required=True)
            command.add_argument("--artifact-digest", type=canonical_digest, required=True)
        if operation in {"publish", "verify"}:
            command.add_argument("--receipt", type=Path, required=True)
        if operation == "publish":
            command.add_argument("--notes", type=Path, required=True)
            command.add_argument("--resume-artifact-id", type=int)
            command.add_argument("--resume-artifact-digest", type=canonical_digest)
    args = parser.parse_args(argv)
    try:
        if args.operation == "assemble":
            assemble(
                args.builds,
                args.directory,
                repo=args.repository,
                version=args.version,
                commit=args.commit,
                run_id=args.run_id,
                attempt=args.attempt,
                source=args.source,
            )
        elif args.operation == "fetch":
            fetch_candidate(
                GitHub(args.repository),
                output=args.directory,
                artifact_id=args.artifact_id,
                digest=args.artifact_digest,
                repository=args.repository,
                version=args.version,
                commit=args.commit,
            )
        else:
            api = GitHub(args.repository)
            if args.operation == "publish" and args.resume_artifact_id is not None:
                require(args.resume_artifact_digest is not None, "resume artifact digest is required")
                restore_progress(
                    api,
                    artifact_id=args.resume_artifact_id,
                    digest=args.resume_artifact_digest,
                    receipt=args.receipt,
                    commit=args.commit,
                )
            session = ReleaseSession(
                api,
                args.directory,
                args.receipt,
                repository=args.repository,
                version=args.version,
                commit=args.commit,
                artifact_id=args.artifact_id,
                artifact_digest=args.artifact_digest,
            )
            if args.operation == "publish":
                result = session.publish(notes=release_notes(args.notes, args.version))
            else:
                session.verify_source()
                session.verify_provenance()
                release = session.load_release(wait=True)
                session.save(releaseId=release["id"])
                result = session.verify_published()
            print(json.dumps(result, ensure_ascii=False))
    except (PublicationError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"publication failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
