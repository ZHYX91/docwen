"""The required source checks shared by local QA and CI/release workflows."""

from __future__ import annotations

import json
import os
import sys
from typing import Any


def architecture_steps() -> list[tuple[str, list[str], bool]]:
    return [
        ("test-governance", [sys.executable, "tools/check_test_governance_consistency.py"], True),
        ("import-linter", [sys.executable, "tools/run_import_linter.py", "--no-cache"], True),
        (
            "architecture-cleanliness",
            [sys.executable, "tools/validation/check_architecture_cleanliness.py", "--repo-root", "."],
            True,
        ),
    ]


def typecheck_steps() -> list[tuple[str, list[str], bool]]:
    # Check every supported target even when QA itself runs on Windows.
    return [
        (
            f"pyright-{platform.lower()}",
            [
                sys.executable,
                "-m",
                "pyright",
                "--level",
                "error",
                "--pythonpath",
                sys.executable,
                "--pythonplatform",
                platform,
            ],
            True,
        )
        for platform in ("Windows", "Linux", "Darwin")
    ]


def validate_ci_results(results: dict[str, Any], event: str) -> None:
    required = {"source-checks"}
    if event == "pull_request":
        required.update({"pytest_windows_pr_integration", "pytest_ubuntu", "pytest_macos", "coverage", "coverage_gui"})
    elif event in {"push", "workflow_dispatch"}:
        required.add("pytest_windows_push")
    else:
        raise ValueError("unsupported source-check event")
    missing = sorted(name for name in required if results.get(name, {}).get("result") != "success")
    if missing:
        raise ValueError(f"required checks failed, missing or skipped: {', '.join(missing)}")


if __name__ == "__main__":
    validate_ci_results(json.loads(os.environ["DOCWEN_CI_RESULTS"]), os.environ["GITHUB_EVENT_NAME"])
