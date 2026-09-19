from __future__ import annotations

import subprocess
from pathlib import Path

import pytest
from scripts.release.publication_contract import PublicationError
from scripts.release.publication_entry import main, route

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "event,operation,ref,artifact,resume,build,publish,verify_hosted",
    [
        ("push", "", "refs/tags/0.12.0", "", "", True, True, False),
        ("workflow_dispatch", "verify", "refs/heads/feature/pr", "", "", True, False, False),
        ("workflow_dispatch", "publish", "refs/heads/feature/pr", "20", "", False, True, False),
        ("workflow_dispatch", "publish", "refs/tags/0.12.0", "20", "80", False, True, False),
        ("workflow_dispatch", "verify", "refs/tags/0.12.0", "20", "", False, False, True),
        ("workflow_dispatch", "publish", "refs/heads/main", "", "", True, True, False),
    ],
)
def test_routes_keep_manual_candidate_acceptance_separate_from_publication(
    event, operation, ref, artifact, resume, build, publish, verify_hosted
) -> None:
    result = route(
        version="0.12.0", source_ref=ref, event=event, operation=operation, artifact_id=artifact, resume_id=resume
    )
    assert result == {
        "version": "0.12.0",
        "source_ref": ref,
        "build": str(build).lower(),
        "publish": str(publish).lower(),
        "verify_hosted": str(verify_hosted).lower(),
    }


@pytest.mark.parametrize(
    "updates",
    [
        {"source_ref": "refs/tags/v0.12.0"},
        {"source_ref": "refs/tags/0.11.0"},
        {"source_ref": "refs/heads/"},
        {"event": "pull_request"},
        {"event": "push", "source_ref": "refs/heads/main"},
        {"operation": "preflight"},
        {"artifact_id": "0"},
        {"artifact_id": "latest"},
        {"resume_id": "80"},
        {"resume_id": "80", "artifact_id": "20"},
        {"version": "0.12.0-rc1"},
    ],
)
def test_routes_reject_ambiguous_or_unrelated_inputs(updates) -> None:
    values = {
        "version": "0.12.0",
        "source_ref": "refs/heads/main",
        "event": "workflow_dispatch",
        "operation": "verify",
        "artifact_id": "",
        "resume_id": "",
    }
    with pytest.raises(PublicationError):
        route(**(values | updates))


@pytest.mark.parametrize("mismatch", [None, "commit", "ref"])
def test_entry_binds_actual_git_checkout_before_writing_job_outputs(tmp_path: Path, mismatch: str | None) -> None:
    source = tmp_path / "source"
    source.mkdir()
    subprocess.run(["git", "init", "-b", "candidate"], cwd=source, check=True, capture_output=True)
    (source / "pyproject.toml").write_text('[project]\nversion = "0.12.0"\n', encoding="utf-8")
    subprocess.run(["git", "add", "pyproject.toml"], cwd=source, check=True, capture_output=True)
    subprocess.run(
        [
            "git",
            "-c",
            "user.name=Fixture",
            "-c",
            "user.email=fixture@example.invalid",
            "-c",
            "commit.gpgsign=false",
            "commit",
            "-m",
            "fixture",
        ],
        cwd=source,
        check=True,
        capture_output=True,
    )
    commit = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=source, text=True).strip()
    output = tmp_path / "outputs"
    output.write_text("previous=value\n", encoding="utf-8")
    result = main(
        [
            "--source",
            str(source),
            "--source-ref",
            "refs/heads/missing" if mismatch == "ref" else "refs/heads/candidate",
            "--commit",
            "f" * 40 if mismatch == "commit" else commit,
            "--event",
            "workflow_dispatch",
            "--github-output",
            str(output),
        ]
    )
    if mismatch:
        assert result == 1
        assert output.read_text(encoding="utf-8") == "previous=value\n"
    else:
        assert result == 0
        assert output.read_text(encoding="utf-8") == (
            "previous=value\nversion=0.12.0\nsource_ref=refs/heads/candidate\n"
            "build=true\npublish=false\nverify_hosted=false\n"
        )
