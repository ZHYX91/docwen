"""Bind a successful hosted source check to an opt-in package-only build."""

from __future__ import annotations

import argparse
import http.client
import json
import os
import re
import subprocess
import tempfile
import zipfile
import zlib
from pathlib import Path
from typing import Any

from scripts.release.publication_contract import PublicationError, checked_lock_sha256, file_identity, require
from scripts.release.publication_download import download_artifact
from scripts.release.publication_http import GitHub

WORKFLOW = ".github/workflows/tests.yml"
RECEIPT = "source-checks.json"
REQUIRED_JOBS = {
    "source-checks": "ubuntu-24.04",
    "Tests on windows-latest": "windows-latest",
    "Tests on ubuntu-24.04": "ubuntu-24.04",
    "Tests on macos-15": "macos-15",
    "Required checks": "ubuntu-24.04",
}


def source_identity(repo: Path) -> dict[str, str]:
    tree = subprocess.check_output(["git", "rev-parse", "HEAD^{tree}"], cwd=repo, text=True, timeout=10).strip()
    return {
        "tree": tree,
        "lockSha256": checked_lock_sha256(repo),
        "workflowSha256": str(file_identity(repo / WORKFLOW)["sha256"]),
    }


def record(path: Path) -> None:
    repo = Path.cwd()
    commit = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=repo, text=True, timeout=10).strip()
    require(commit == os.environ["GITHUB_SHA"], "source check checkout differs from the hosted event")
    with path.open("x", encoding="utf-8") as stream:
        json.dump(
            {
                "schema": "docwen.source-checks/v1",
                "repository": os.environ["GITHUB_REPOSITORY"],
                "runId": int(os.environ["GITHUB_RUN_ID"]),
                "attempt": int(os.environ["GITHUB_RUN_ATTEMPT"]),
                "event": os.environ["GITHUB_EVENT_NAME"],
                "commit": commit,
                **source_identity(repo),
            },
            stream,
        )


def verify_run(run: dict[str, Any], jobs: list[dict[str, Any]], *, repository: str, run_id: int) -> None:
    require(
        run["id"] == run_id
        and run["path"] == WORKFLOW
        and run["event"] in {"pull_request", "push", "workflow_dispatch"}
        and run["repository"]["full_name"] == run["head_repository"]["full_name"] == repository
        and run["status"] == "completed"
        and run["conclusion"] == "success",
        "use a successful same-repository tests workflow",
    )
    require(
        len(jobs) == len(REQUIRED_JOBS)
        and {job["name"]: job["labels"] for job in jobs} == {name: [label] for name, label in REQUIRED_JOBS.items()}
        and all(
            job["run_id"] == run_id
            and job["run_attempt"] == run["run_attempt"]
            and job["head_sha"] == run["head_sha"]
            and job["status"] == "completed"
            and job["conclusion"] == "success"
            for job in jobs
        ),
        "source-check job graph, runner labels or outcomes differ",
    )


def verify_receipt(
    receipt: dict[str, Any], tested: dict[str, Any], *, expected: dict[str, str], run: dict[str, Any]
) -> None:
    require(
        receipt["schema"] == "docwen.source-checks/v1"
        and receipt["repository"] == run["repository"]["full_name"]
        and receipt["runId"] == run["id"]
        and receipt["attempt"] == run["run_attempt"]
        and receipt["event"] == run["event"]
        and receipt["commit"] == tested["sha"]
        and receipt["tree"] == tested["tree"]["sha"],
        "source-check receipt is bound to a different run or checkout",
    )
    if run["event"] == "pull_request":
        parents = tested["parents"]
        require(len(parents) == 2 and parents[1]["sha"] == run["head_sha"], "PR merge checkout differs")
    else:
        require(tested["sha"] == run["head_sha"], "tested commit differs")
    require(all(receipt[name] == value for name, value in expected.items()), "source tree, lock or workflow changed")


def authorize(api: GitHub, run_id: int) -> None:
    prefix = f"/repos/{api.repository}"
    run_path = f"{prefix}/actions/runs/{run_id}"
    run = api.request("GET", run_path, timeout=6)
    attempt = run["run_attempt"]
    require(type(attempt) is int and attempt > 0, "invalid source-check attempt")
    jobs = api.request("GET", f"{run_path}/attempts/{attempt}/jobs", timeout=6)["jobs"]
    verify_run(run, jobs, repository=api.repository, run_id=run_id)
    name = f"source-checks-{run_id}-{attempt}"
    artifacts = api.request("GET", f"{run_path}/artifacts?name={name}", timeout=6)["artifacts"]
    require(len(artifacts) == 1, "exact source-check receipt missing or ambiguous")
    artifact = artifacts[0]
    require(
        artifact["name"] == name
        and artifact["expired"] is False
        and artifact["workflow_run"]["id"] == run_id
        and type(artifact["size_in_bytes"]) is int
        and 0 < artifact["size_in_bytes"] <= 16_384,
        "source-check receipt artifact unavailable",
    )
    with tempfile.TemporaryDirectory(prefix="docwen-source-checks-") as temporary:
        archive = Path(temporary) / "proof.zip"
        download_artifact(api, artifact, archive)
        with zipfile.ZipFile(archive) as proof:
            entries = proof.infolist()
            require(
                len(entries) == 1
                and entries[0].filename == RECEIPT
                and entries[0].file_size <= 16_384
                and not entries[0].flag_bits & 1
                and entries[0].compress_type in {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED},
                "invalid receipt ZIP",
            )
            receipt = json.loads(proof.read(entries[0]))
    commit = receipt["commit"]
    require(isinstance(commit, str) and re.fullmatch(r"[0-9a-f]{40}", commit) is not None, "invalid tested commit")
    tested = api.request("GET", f"{prefix}/git/commits/{commit}", timeout=6)
    verify_receipt(receipt, tested, expected=source_identity(Path.cwd()), run=run)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="action", required=True)
    create = sub.add_parser("record")
    create.add_argument("output", type=Path)
    reuse = sub.add_parser("reuse")
    reuse.add_argument("--run-id", required=True)
    reuse.add_argument("--github-output", type=Path, required=True)
    args = parser.parse_args()
    if args.action == "record":
        record(args.output)
        return
    reused = False
    try:
        require(re.fullmatch(r"[1-9][0-9]{0,17}", args.run_id) is not None, "positive source-check run ID required")
        run_id = int(args.run_id)
        authorize(GitHub(os.environ["GITHUB_REPOSITORY"]), run_id)
    except (
        PublicationError,
        OSError,
        ValueError,
        KeyError,
        TypeError,
        http.client.HTTPException,
        subprocess.SubprocessError,
        zipfile.BadZipFile,
        zlib.error,
    ):
        print("Source checks could not be reused. Running the full source checks.")
    else:
        reused = True
        print(f"Reusing source checks {run_id} for package-only verification.")
    with args.github_output.open("a", encoding="utf-8") as stream:
        stream.write(f"reused={str(reused).lower()}\n")


if __name__ == "__main__":
    main()
