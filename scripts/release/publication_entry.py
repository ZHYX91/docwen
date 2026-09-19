"""Resolve one workflow route from its selected Git ref and source version."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tomllib
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.release.publication_contract import PublicationError, package_names, require


def route(
    *, version: str, source_ref: str, event: str, operation: str, artifact_id: str, resume_id: str
) -> dict[str, str]:
    package_names(version)
    require(event in {"push", "workflow_dispatch"}, "unsupported publication trigger")
    for identity in (artifact_id, resume_id):
        require(not identity or re.fullmatch(r"[1-9]\d*", identity) is not None, "positive artifact ID required")
    if event == "push":
        require(source_ref == f"refs/tags/{version}", "numeric tag must match the source version")
        require(not artifact_id and not resume_id, "tag push cannot select recovery artifacts")
        operation = "publish"
    else:
        require(operation in {"verify", "publish"}, "unknown publication operation")
        require(
            source_ref == f"refs/tags/{version}"
            or (source_ref.startswith("refs/heads/") and len(source_ref) > len("refs/heads/")),
            "select a source branch or its matching numeric tag",
        )
    require(
        not resume_id or (bool(artifact_id) and operation == "publish"),
        "resume requires publish and an exact candidate ID",
    )
    return {
        "version": version,
        "source_ref": source_ref,
        "build": str(not artifact_id).lower(),
        "publish": str(operation == "publish").lower(),
        "verify_hosted": str(operation == "verify" and bool(artifact_id)).lower(),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path.cwd())
    parser.add_argument("--source-ref", required=True)
    parser.add_argument("--commit", required=True)
    parser.add_argument("--event", required=True)
    parser.add_argument("--operation", default="verify")
    parser.add_argument("--artifact-id", default="")
    parser.add_argument("--resume-id", default="")
    parser.add_argument("--github-output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        source = args.source.resolve(strict=True)
        version = tomllib.loads((source / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
        result = route(
            version=version,
            source_ref=args.source_ref,
            event=args.event,
            operation=args.operation,
            artifact_id=args.artifact_id,
            resume_id=args.resume_id,
        )
        require(re.fullmatch(r"[0-9a-f]{40}", args.commit) is not None, "invalid source commit")
        for ref in ("HEAD", f"{args.source_ref}^{{commit}}"):
            actual = subprocess.check_output(
                ["git", "rev-parse", "--verify", ref], cwd=source, text=True, timeout=30
            ).strip()
            require(actual == args.commit, "checkout and selected source ref differ")
        with args.github_output.open("a", encoding="utf-8") as stream:
            for name, value in result.items():
                stream.write(f"{name}={value}\n")
        print(json.dumps(result))
    except (PublicationError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"publication entry failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
