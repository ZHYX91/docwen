from __future__ import annotations

import hashlib
import io
import subprocess
import zipfile
from pathlib import Path

import pytest
from scripts.release import publication
from scripts.release.publication_contract import PublicationError, canonical_digest, canonical_json
from tests.support.publication import COMMIT, REPOSITORY

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("failure", [None, "active", "archive", "source", "inventory"])
def test_progress_restore_requires_completed_owner_and_exact_archive(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, failure: str | None
) -> None:
    state = {"sourceCommit": "f" * 40 if failure == "source" else COMMIT, "pending": "publish"}
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("other.json" if failure == "inventory" else "publication-progress.json", canonical_json(state))
    content = buffer.getvalue()
    digest = "sha256:" + hashlib.sha256(content).hexdigest()

    class API:
        repository = REPOSITORY

        def get(self, path):
            if "/artifacts/" in path:
                return {
                    "id": 80,
                    "expired": False,
                    "digest": digest,
                    "name": "docwen-publication-progress-70-1",
                    "workflow_run": {"id": 70},
                }
            return {
                "status": "in_progress" if failure == "active" else "completed",
                "head_sha": COMMIT,
                "path": ".github/workflows/release.yml",
            }

    def download(*args, **kwargs):
        return subprocess.CompletedProcess(args[0], 0, stdout=content + (b"changed" if failure == "archive" else b""))

    monkeypatch.setattr(publication.subprocess, "run", download)
    receipt = tmp_path / "progress.json"
    if failure:
        with pytest.raises(PublicationError):
            publication.restore_progress(API(), artifact_id=80, digest=digest, receipt=receipt, commit=COMMIT)
        assert not receipt.exists()
    else:
        publication.restore_progress(API(), artifact_id=80, digest=digest, receipt=receipt, commit=COMMIT)
        assert receipt.read_bytes() == canonical_json(state)
        publication.restore_progress(API(), artifact_id=80, digest=digest, receipt=receipt, commit=COMMIT)


def test_digest_normalization_and_exact_release_notes(tmp_path: Path) -> None:
    digest = "a" * 64
    assert canonical_digest(digest) == canonical_digest("sha256:" + digest)
    with pytest.raises(PublicationError):
        canonical_digest("short")
    notes = tmp_path / "CHANGELOG.md"
    notes.write_text("# Changes\n\n## 0.10.0\nCurrent\n\n## 0.9.1\nOld\n", encoding="utf-8")
    assert publication.release_notes(notes, "0.10.0") == "## 0.10.0\nCurrent"
    with pytest.raises(PublicationError):
        publication.release_notes(notes, "0.10")
