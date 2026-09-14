from __future__ import annotations

import pytest
from scripts.release.publication_contract import PublicationError
from scripts.release.resolve_publication import resolve
from tests.support.publication import COMMIT, REPOSITORY

pytestmark = pytest.mark.unit

DIGEST = "sha256:" + "d" * 64
RUNS_PATH = f"/repos/{REPOSITORY}/actions/runs?head_sha={COMMIT}&status=success&per_page=100"


def _run(
    run_id: int,
    *,
    attempt: int = 1,
    conclusion: str = "success",
    path: str = ".github/workflows/release.yml",
    head_sha: str = COMMIT,
) -> dict:
    return {
        "id": run_id,
        "run_attempt": attempt,
        "head_sha": head_sha,
        "conclusion": conclusion,
        "path": path,
    }


def _artifact(
    name: str,
    *,
    artifact_id: int,
    run_id: int,
    expired: bool = False,
    digest: str | None = DIGEST,
    head_sha: str = COMMIT,
) -> dict:
    return {
        "id": artifact_id,
        "name": name,
        "expired": expired,
        "digest": digest,
        "workflow_run": {"id": run_id, "head_sha": head_sha},
    }


class API:
    repository = REPOSITORY

    def __init__(self, responses: dict[str, dict]) -> None:
        self.responses = responses
        self.paths: list[str] = []

    def get(self, path: str, **_kwargs: object) -> dict:
        self.paths.append(path)
        return self.responses[path]


def test_resolver_selects_the_newest_exact_preflight_artifact() -> None:
    api = API(
        {
            RUNS_PATH: {"workflow_runs": [_run(70), _run(90)]},
            f"/repos/{REPOSITORY}/actions/runs/90/artifacts": {
                "artifacts": [_artifact("docwen-publication-90-1", artifact_id=900, run_id=90)]
            },
        }
    )

    assert resolve(api, COMMIT) == {
        "artifactId": 900,
        "artifactDigest": DIGEST,
        "runId": 90,
        "runAttempt": 1,
        "name": "docwen-publication-90-1",
    }
    assert f"/repos/{REPOSITORY}/actions/runs/70/artifacts" not in api.paths


@pytest.mark.parametrize(
    "artifacts",
    [
        [_artifact("docwen-publication-90-1", artifact_id=900, run_id=90, expired=True)],
        [_artifact("docwen-publication-90-1", artifact_id=900, run_id=91)],
        [_artifact("docwen-publication-90-1", artifact_id=900, run_id=90, head_sha="e" * 40)],
        [_artifact("docwen-publication-90-1", artifact_id=900, run_id=90, digest=None)],
        [_artifact("docwen-publication-progress-90-1", artifact_id=900, run_id=90)],
        [],
    ],
)
def test_resolver_fails_closed_on_a_foreign_or_unusable_artifact(artifacts: list[dict]) -> None:
    api = API(
        {
            RUNS_PATH: {"workflow_runs": [_run(90)]},
            f"/repos/{REPOSITORY}/actions/runs/90/artifacts": {"artifacts": artifacts},
        }
    )

    with pytest.raises(PublicationError):
        resolve(api, COMMIT)


@pytest.mark.parametrize(
    "runs",
    [
        [],
        [_run(90, conclusion="failure")],
        [_run(90, path=".github/workflows/tests.yml")],
        [_run(90, head_sha="f" * 40)],
    ],
)
def test_resolver_requires_a_successful_release_workflow_run_for_the_commit(runs: list[dict]) -> None:
    api = API({RUNS_PATH: {"workflow_runs": runs}})

    with pytest.raises(PublicationError):
        resolve(api, COMMIT)
