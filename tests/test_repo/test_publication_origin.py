from __future__ import annotations

import hashlib
import io
import json
import subprocess
import zipfile
from pathlib import Path

import pytest
from scripts.release import publication
from scripts.release.publication_contract import PublicationError, verify_origin
from scripts.release.publication_session import ReleaseSession, verify_candidate_jobs
from tests.support.publication import COMMIT, DIGEST, REPOSITORY, VERSION, FakeGitHub, candidate

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "status,conclusion", [("in_progress", None), ("completed", "success"), ("completed", "failure")]
)
@pytest.mark.parametrize("event", ["push", "workflow_dispatch"])
def test_candidate_origin_accepts_its_active_run_and_recovery_after_publisher_failure(
    tmp_path: Path, status: str, conclusion: str | None, event: str
) -> None:
    _, manifest = candidate(tmp_path)
    api = FakeGitHub()
    api.run.update(status=status, conclusion=conclusion, event=event)
    verify_origin(manifest, api.run, api.artifact, digest=DIGEST)
    verify_candidate_jobs(api.jobs)


@pytest.mark.parametrize(
    "field,value",
    [
        ("run_attempt", 2),
        ("head_sha", "f" * 40),
        ("head_branch", "main"),
        ("event", "pull_request"),
        ("status", "queued"),
    ],
)
def test_candidate_origin_rejects_another_attempt_source_ref_or_trigger(tmp_path: Path, field: str, value) -> None:
    _, manifest = candidate(tmp_path)
    api = FakeGitHub()
    api.run[field] = value
    with pytest.raises(PublicationError):
        verify_origin(manifest, api.run, api.artifact, digest=DIGEST)


@pytest.mark.parametrize("failure", ["failure", "skipped", "cancelled", "missing", "duplicate"])
def test_candidate_requires_every_producer_job_even_when_the_parent_run_is_complete(failure: str) -> None:
    api = FakeGitHub()
    if failure == "missing":
        api.jobs["jobs"].pop()
    elif failure == "duplicate":
        api.jobs["jobs"].append(api.jobs["jobs"][-1].copy())
    else:
        api.jobs["jobs"][-1]["conclusion"] = failure
    api.jobs["total_count"] = len(api.jobs["jobs"])
    with pytest.raises(PublicationError, match="required candidate job did not pass"):
        verify_candidate_jobs(api.jobs)


@pytest.mark.parametrize("branch_candidate", [False, True])
def test_fetch_publish_and_independent_readback_reuse_one_exact_candidate(
    tmp_path: Path, monkeypatch, branch_candidate: bool
) -> None:
    """Exercise all CLI stages with a real ZIP and in-memory remote writes."""

    source_ref = "refs/heads/release/candidate" if branch_candidate else f"refs/tags/{VERSION}"
    directory, _ = candidate(tmp_path, source_ref=source_ref)
    archive = io.BytesIO()
    with zipfile.ZipFile(archive, "w") as bundle:
        for path in directory.iterdir():
            bundle.writestr(path.name, path.read_bytes())
    content = archive.getvalue()
    api = FakeGitHub()
    api.run.update(
        status="in_progress",
        conclusion=None,
        event="workflow_dispatch" if branch_candidate else "push",
        head_branch=source_ref.split("/", 2)[2],
    )
    if branch_candidate:
        api.tag = None
    api.artifact["digest"] = "sha256:" + hashlib.sha256(content).hexdigest()
    monkeypatch.setattr(publication, "GitHub", lambda _repository: api)
    provenance_commands = []

    def run(args, **kwargs):
        if args[:2] == ["gh", "api"]:
            assert args[2] == f"/repos/{REPOSITORY}/actions/artifacts/20/zip"
            kwargs["stdout"].write(content)
        else:
            assert args[:3] == ["gh", "attestation", "verify"]
            provenance_commands.append(args)
        return subprocess.CompletedProcess(args, 0, stdout="", stderr="")

    monkeypatch.setattr(publication.subprocess, "run", run)
    fetched = tmp_path / "fetched"
    common = [
        "--repository",
        REPOSITORY,
        "--version",
        VERSION,
        "--commit",
        COMMIT,
        "--artifact-id",
        "20",
        "--directory",
        str(fetched),
    ]
    assert publication.main(["fetch", *common]) == 0
    assert {p.name: p.read_bytes() for p in fetched.iterdir()} == {p.name: p.read_bytes() for p in directory.iterdir()}
    inspection = tmp_path / "inspection.json"
    assert publication.main(["inspect", *common, "--receipt", str(inspection)]) == 0
    inspected = json.loads(inspection.read_text(encoding="utf-8"))
    assert inspected["stage"] == "candidate-verified"
    assert inspected["origin"]["sourceRef"] == source_ref
    assert inspected["provenance"] == "verified"
    assert not api.writes and not api.downloads
    if branch_candidate:
        assert api.tag is None
    notes = tmp_path / "CHANGELOG.md"
    notes.write_text(f"# Changes\n\n## {VERSION}\nReady\n", encoding="utf-8")
    receipt = tmp_path / "progress.json"
    assert publication.main(["publish", *common, "--receipt", str(receipt), "--notes", str(notes)]) == 0
    assert len(api.writes) == (7 if branch_candidate else 6)
    assert not api.downloads
    writes = api.writes[:]
    assert publication.main(["verify", *common, "--receipt", str(tmp_path / "readback.json")]) == 0
    assert api.writes == writes
    assert len(api.downloads) == 4
    assert len(provenance_commands) == 15
    assert all(command[command.index("--source-ref") + 1] == source_ref for command in provenance_commands)
    run_reads = [path for path in api.reads if "/actions/runs/" in path and "/jobs?" not in path]
    assert run_reads and set(run_reads) == {f"/repos/{REPOSITORY}/actions/runs/10/attempts/1"}


@pytest.mark.parametrize("field,value", [("id", 21), ("expired", True), ("digest", None), ("digest", "short")])
def test_explicit_artifact_lookup_rejects_missing_or_unusable_identity(field, value) -> None:
    api = FakeGitHub()
    api.artifact[field] = value
    with pytest.raises(PublicationError):
        publication.artifact_digest(api, 20)
    assert api.writes == []


def test_candidate_source_verification_reads_original_attempt_after_workflow_rerun(tmp_path: Path) -> None:
    directory, _ = candidate(tmp_path)
    api = FakeGitHub()
    session = ReleaseSession(
        api,
        directory,
        tmp_path / "progress.json",
        repository=REPOSITORY,
        version=VERSION,
        commit=COMMIT,
        artifact_id=20,
        artifact_digest=DIGEST,
    )
    session.verify_source()
    assert f"/repos/{REPOSITORY}/actions/runs/10/attempts/1" in api.reads
    assert f"/repos/{REPOSITORY}/actions/runs/10" not in api.reads
