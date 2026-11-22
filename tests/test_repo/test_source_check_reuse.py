import json
import sys
import zipfile

import pytest
from scripts.release import source_check_reuse as reuse

pytestmark = pytest.mark.unit
RUN_ID = 17
REPO = "owner/docwen"
HEAD = "a" * 40
MERGE = "b" * 40
IDENTITY = {"tree": "c" * 40, "lockSha256": "d" * 64, "workflowSha256": "e" * 64}


@pytest.mark.parametrize(
    "case", ["match", "tree", "lock", "workflow", "missing", "wrong-attempt", "fork", "failed-job", "label"]
)
def test_source_check_reuse_admits_only_exact_package_checks(tmp_path, monkeypatch, case):
    origin = {"full_name": REPO}
    run = {
        "id": RUN_ID,
        "path": reuse.WORKFLOW,
        "event": "pull_request",
        "repository": origin,
        "head_repository": origin if case != "fork" else {"full_name": "fork/docwen"},
        "status": "completed",
        "conclusion": "success",
        "head_sha": HEAD,
        "run_attempt": 2,
    }
    artifact = {
        "id": 97,
        "name": "source-checks-17-2",
        "size_in_bytes": 600,
        "expired": False,
        "workflow_run": {"id": RUN_ID},
    }
    jobs = [
        {
            "name": name,
            "labels": [runner],
            "run_id": RUN_ID,
            "run_attempt": 2,
            "head_sha": HEAD,
            "status": "completed",
            "conclusion": "success",
        }
        for name, runner in reuse.REQUIRED_JOBS.items()
    ]
    if case == "failed-job":
        jobs[1]["conclusion"] = "failure"
    if case == "label":
        jobs[1]["labels"] = ["self-hosted"]
    source = {"sha": MERGE, "parents": [{"sha": "1" * 40}, {"sha": HEAD}], "tree": {"sha": IDENTITY["tree"]}}
    receipt = {
        "schema": "docwen.source-checks/v1",
        "commit": MERGE,
        "runId": RUN_ID,
        "attempt": 1 if case == "wrong-attempt" else 2,
        "event": "pull_request",
        "repository": REPO,
        **IDENTITY,
    }
    if case == "lock":
        receipt["lockSha256"] = "0" * 64
    elif case == "workflow":
        receipt["workflowSha256"] = "0" * 64
    elif case == "tree":
        receipt["tree"] = source["tree"]["sha"] = "0" * 40

    class FakeGitHub:
        repository = REPO

        def request(self, method, path, *, timeout):
            assert method == "GET"
            prefix = f"/repos/{REPO}"
            run_path = f"{prefix}/actions/runs/{RUN_ID}"
            return {
                run_path: run,
                f"{run_path}/attempts/2/jobs": {"jobs": jobs},
                f"{run_path}/artifacts?name=source-checks-17-2": {"artifacts": [] if case == "missing" else [artifact]},
                f"{prefix}/git/commits/{MERGE}": source,
            }[path]

    def download(_api, selected, archive):
        assert selected == artifact
        with zipfile.ZipFile(archive, "w") as bundle:
            bundle.writestr(reuse.RECEIPT, json.dumps(receipt))

    output = tmp_path / "output"
    monkeypatch.setenv("GITHUB_REPOSITORY", REPO)
    monkeypatch.setattr(reuse, "GitHub", lambda _repository: FakeGitHub())
    monkeypatch.setattr(reuse, "download_artifact", download)
    monkeypatch.setattr(reuse, "source_identity", lambda _repo: IDENTITY)
    monkeypatch.setattr(
        sys, "argv", ["source_check_reuse", "reuse", "--run-id", str(RUN_ID), "--github-output", str(output)]
    )
    reuse.main()
    assert output.read_text(encoding="utf-8").strip() == f"reused={str(case == 'match').lower()}"
