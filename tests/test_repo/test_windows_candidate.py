"""Independent Windows acceptance keeps provenance and full publication separate."""

from __future__ import annotations

import copy
import io
import json
import zipfile
from pathlib import Path

import pytest
from scripts.release import publication, publication_download, publication_platform, windows_candidate
from scripts.release.publication_contract import PublicationError, assemble, verify_inventory
from scripts.release.publication_session import candidate_jobs, verify_candidate_jobs
from tests.support.publication import COMMIT, REPOSITORY, VERSION, FakeGitHub, candidate
from tests.test_repo.test_publication_platform import PlatformGitHub

pytestmark = pytest.mark.unit


@pytest.fixture
def producer(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    directory, _ = candidate(tmp_path)
    windows_candidate.create(
        directory,
        repository=REPOSITORY,
        version=VERSION,
        commit=COMMIT,
        run_id=10,
        attempt=1,
        source_ref=f"refs/tags/{VERSION}",
    )
    api = PlatformGitHub(directory)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.write(directory / windows_candidate.NAME, windows_candidate.NAME)
    api.set_archive(25, buffer.getvalue(), "windows-metadata")
    api.artifact = next(copy.deepcopy(row) for row in api.platform_artifacts if row["id"] == 22)
    for job in api.jobs["jobs"]:
        if job["name"] in {"Build Linux packages", "verify-release"}:
            job["conclusion"] = "failure"
    api.run["conclusion"] = "failure"
    monkeypatch.setattr(publication.subprocess, "run", api.execute)
    monkeypatch.setattr(publication_download, "open_response", api.open_response)
    monkeypatch.setattr(publication_platform, "download_artifact", publication_download.download_artifact_in_process)
    monkeypatch.setattr(publication, "GitHub", lambda _: api)
    return api, directory, tmp_path / "selected", tmp_path / "receipt.json"


def fetch(producer):
    api, _, output, receipt = producer
    return windows_candidate.fetch(
        api,
        output=output,
        receipt=receipt,
        artifact_id=22,
        digest=api.artifact["digest"],
        repository=REPOSITORY,
        version=VERSION,
        commit=COMMIT,
    )


def test_windows_can_be_accepted_while_linux_build_failed_but_cannot_be_published(producer) -> None:
    api, _, output, receipt = producer
    result = fetch(producer)
    assert result["candidateKind"] == "windows-producer"
    assert result["assets"].keys() == {windows_candidate.ARCHIVE}
    assert set(api.downloads) == {"22", "25"}
    assert set(api.provenance) == {windows_candidate.ARCHIVE, windows_candidate.NAME}
    assert json.loads(receipt.read_text())["provenance"] == "verified"
    with pytest.raises((PublicationError, FileNotFoundError)):
        verify_inventory(output, repository=REPOSITORY, version=VERSION, commit=COMMIT)
    with pytest.raises(PublicationError):
        verify_candidate_jobs(api.jobs)
    assert not api.writes
    downloads = list(api.downloads)
    assert fetch(producer) == result
    assert api.downloads == downloads


def test_public_fetch_platform_dispatches_exact_windows_producer(producer) -> None:
    api, _, output, receipt = producer
    assert (
        publication.main(
            [
                "fetch-platform",
                "--repository",
                REPOSITORY,
                "--version",
                VERSION,
                "--commit",
                COMMIT,
                "--artifact-id",
                "22",
                "--platform",
                "windows",
                "--directory",
                str(output),
                "--receipt",
                str(receipt),
            ]
        )
        == 0
    )
    assert json.loads(receipt.read_text())["candidateKind"] == "windows-producer"
    assert not api.writes


@pytest.mark.parametrize(
    "failure",
    [
        "source-check",
        "windows",
        "sha",
        "workflow",
        "fork",
        "attempt",
        "metadata-missing",
        "attestation",
        "digest",
        "archive",
    ],
)
def test_independent_windows_rejects_unverified_origin_or_content(producer, failure) -> None:
    api, _, _, receipt = producer
    if failure in {"source-check", "windows"}:
        name = "source-checks" if failure == "source-check" else "Build Windows packages"
        next(job for job in api.jobs["jobs"] if job["name"] == name)["conclusion"] = "failure"
    elif failure == "sha":
        api.run["head_sha"] = "f" * 40
    elif failure == "workflow":
        api.run["path"] = ".github/workflows/other.yml"
    elif failure == "fork":
        api.run["repository"] = {"full_name": "fork/repo"}
    elif failure == "attempt":
        api.run["run_attempt"] = 2
    elif failure == "metadata-missing":
        api.platform_artifacts[:] = [row for row in api.platform_artifacts if row["id"] != 25]
    elif failure == "attestation":
        api.fail_provenance = True
    elif failure == "digest":
        api.archives[22] += b"changed"
    elif failure == "archive":
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr(windows_candidate.ARCHIVE, b"changed")
        api.set_archive(22, buffer.getvalue(), "windows")
        api.artifact = next(row for row in api.platform_artifacts if row["id"] == 22)
    with pytest.raises(PublicationError):
        fetch(producer)
    assert not receipt.exists()


def test_full_assembly_rejects_rebuilt_windows_bytes(producer, tmp_path: Path) -> None:
    _, directory, _, _ = producer
    (tmp_path / "builds/windows" / windows_candidate.ARCHIVE).write_bytes(b"rebuilt")
    with pytest.raises(PublicationError, match="differ"):
        assemble(
            tmp_path / "builds",
            tmp_path / "final",
            repo=REPOSITORY,
            version=VERSION,
            commit=COMMIT,
            run_id=10,
            attempt=2,
            source=tmp_path / "source",
            source_ref=f"refs/tags/{VERSION}",
            windows_manifest=directory / windows_candidate.NAME,
        )
    assert not (tmp_path / "final").exists()


@pytest.mark.parametrize("failure", [None, "new-failure", "different-source", "duplicate", "truncated"])
def test_rerun_only_reuses_same_run_success_and_never_masks_new_failures(failure) -> None:
    class Rerun(FakeGitHub):
        def get(self, path, *, allow_missing=False):
            number = int(path.split("/attempts/")[1].split("/")[0])
            if "/jobs?" not in path:
                return {
                    **self.run,
                    "run_attempt": number,
                    "head_sha": "f" * 40 if failure == "different-source" and number == 1 else COMMIT,
                }
            if number == 1:
                return copy.deepcopy(self.jobs)
            jobs = [
                {"name": "Build Linux packages", "conclusion": "success"},
                {"name": "verify-release", "conclusion": "success"},
            ]
            if failure == "new-failure":
                jobs.append({"name": "Build Windows packages", "conclusion": "failure"})
            if failure == "duplicate":
                jobs.append(jobs[0])
            return {"total_count": len(jobs) + int(failure == "truncated"), "jobs": jobs}

    api = Rerun()
    if failure:
        with pytest.raises(PublicationError):
            verify_candidate_jobs(candidate_jobs(api, run_id=10, attempt=2, commit=COMMIT))
    else:
        jobs = candidate_jobs(api, run_id=10, attempt=2, commit=COMMIT)
        verify_candidate_jobs(jobs)
        assert next(row for row in jobs["jobs"] if row["name"] == "Build Windows packages")["run_attempt"] == 1


@pytest.mark.parametrize("replacement", ["carried", "rebuilt", "failure", "no-timestamps", "changed-steps"])
def test_carried_job_new_id_does_not_move_its_artifact_to_the_rerun(replacement) -> None:
    class CopiedJobs(FakeGitHub):
        def get(self, path, *, allow_missing=False):
            number = int(path.split("/attempts/")[1].split("/")[0])
            if "/jobs?" not in path:
                return {**self.run, "run_attempt": number}
            jobs = copy.deepcopy(self.jobs)
            for row in jobs["jobs"]:
                row.update(
                    id=100 + number,
                    run_attempt=number,
                    started_at="2026-11-12T11:51:27Z",
                    completed_at="2026-11-12T11:58:31Z",
                    steps=[{"number": 1, "name": "Build", "conclusion": "success"}],
                )
                if number == 2:
                    if replacement == "rebuilt":
                        row["started_at"] = "2026-11-12T12:44:00Z"
                    elif replacement == "failure":
                        row["conclusion"] = "failure"
                    elif replacement == "no-timestamps":
                        row.pop("started_at")
                    elif replacement == "changed-steps":
                        row["steps"][0]["name"] = "Different build"
            return jobs

    jobs = candidate_jobs(CopiedJobs(), run_id=10, attempt=2, commit=COMMIT)
    windows = next(row for row in jobs["jobs"] if row["name"] == "Build Windows packages")
    assert windows["run_attempt"] == 2
    assert windows["id"] == 102
    assert windows["producer_attempt"] == (1 if replacement == "carried" else 2)
    if replacement == "failure":
        with pytest.raises(PublicationError):
            verify_candidate_jobs(jobs)
