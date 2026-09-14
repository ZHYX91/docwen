"""Resolve the exact preflight publication artifact owned by one release commit.

Publication never rebuilds: a tag-triggered release must find the artifact that a
preflight run produced for the same commit and fail closed when it is missing,
expired, or belongs to another run.  The resolved identity is fed to
``publication.py fetch``/``publish``, which re-validates it against the platform.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.release.publication_contract import PublicationError, require
from scripts.release.publication_http import GitHub

WORKFLOW = ".github/workflows/release.yml"
ARTIFACT_NAME = re.compile(r"docwen-publication-([1-9]\d*)-([1-9]\d*)")
DIGEST = re.compile(r"sha256:[0-9a-f]{64}")


def successful_preflights(api: GitHub, commit: str, *, limit: int) -> list[dict]:
    """Return successful release-workflow runs for ``commit``, newest first."""

    payload = api.get(f"/repos/{api.repository}/actions/runs?head_sha={commit}&status=success&per_page=100")
    runs = payload.get("workflow_runs")
    require(isinstance(runs, list), "unexpected workflow runs response")
    selected = [
        run
        for run in runs
        if run.get("head_sha") == commit and run.get("conclusion") == "success" and run.get("path") == WORKFLOW
    ]
    selected.sort(key=lambda run: int(run.get("id", 0)), reverse=True)
    require(selected, f"no successful preflight run is recorded for {commit}")
    return selected[:limit]


def resolve(api: GitHub, commit: str, *, limit: int = 10) -> dict[str, object]:
    """Return the preflight publication artifact identity for ``commit``."""

    for run in successful_preflights(api, commit, limit=limit):
        run_id = int(run["id"])
        attempt = int(run.get("run_attempt", 1))
        payload = api.get(f"/repos/{api.repository}/actions/runs/{run_id}/artifacts")
        artifacts = payload.get("artifacts")
        require(isinstance(artifacts, list), "unexpected run artifacts response")
        for artifact in artifacts:
            name = str(artifact.get("name", ""))
            match = ARTIFACT_NAME.fullmatch(name)
            if match is None:
                continue
            owner = artifact.get("workflow_run") or {}
            require(
                (int(match.group(1)), int(match.group(2))) == (run_id, attempt)
                and artifact.get("expired") is False
                and owner.get("id") == run_id
                and owner.get("head_sha") == commit,
                "preflight publication artifact identity mismatch",
            )
            digest = artifact.get("digest")
            require(
                isinstance(digest, str) and DIGEST.fullmatch(digest) is not None,
                "preflight publication artifact digest is missing",
            )
            require(type(artifact.get("id")) is int and artifact["id"] > 0, "artifact ID missing")
            return {
                "artifactId": artifact["id"],
                "artifactDigest": digest,
                "runId": run_id,
                "runAttempt": attempt,
                "name": name,
            }
    raise PublicationError(f"no unexpired preflight publication artifact is recorded for {commit}")


def _append(path: Path, lines: dict[str, object]) -> None:
    with path.open("a", encoding="utf-8") as stream:
        for key, value in lines.items():
            stream.write(f"{key}={value}\n")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository", required=True)
    parser.add_argument("--commit", required=True)
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument("--github-env", type=Path)
    parser.add_argument("--github-output", type=Path)
    args = parser.parse_args(argv)
    try:
        resolved = resolve(GitHub(args.repository), args.commit, limit=args.limit)
    except (PublicationError, OSError, ValueError) as error:
        print(f"preflight artifact resolution failed: {error}", file=sys.stderr)
        return 1
    if args.github_env is not None:
        _append(args.github_env, {"ARTIFACT_ID": resolved["artifactId"], "ARTIFACT_DIGEST": resolved["artifactDigest"]})
    if args.github_output is not None:
        _append(
            args.github_output,
            {"artifact_id": resolved["artifactId"], "artifact_digest": resolved["artifactDigest"]},
        )
    print(json.dumps(resolved, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
