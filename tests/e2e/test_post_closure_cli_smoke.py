"""E2E smoke test: calls the CLI smoke validation runner and asserts success."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.e2e

REPO_ROOT = Path(__file__).resolve().parents[2]
SMOKE_SCRIPT = REPO_ROOT / "tools" / "validation" / "run_post_closure_cli_smoke.py"


def test_post_closure_cli_smoke() -> None:
    """Run the post-closure CLI smoke tool and assert all critical tests pass."""
    import os as _os

    env = _os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = [sys.executable, str(SMOKE_SCRIPT)]
    r = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=600,
        cwd=str(REPO_ROOT),
        env=env,
    )
    # The smoke script prints exactly one JSON object to stdout.
    # Its own subprocess invocations have their stdout/stderr captured
    # separately, so stdout of the smoke script itself is clean JSON.
    stdout = r.stdout.strip()
    try:
        data = json.loads(stdout)
        assert isinstance(data, dict) and "passed" in data and "results" in data
    except (json.JSONDecodeError, ValueError, AssertionError):
        pytest.fail(f"Could not parse smoke JSON (rc={r.returncode}).\nstdout tail:\n{stdout[-3000:]}")

    # Enumerate failures for readable assertion messages.
    failures = [res for res in data["results"] if not res.get("optional") and not res["passed"]]
    if failures:
        names = [f["name"] for f in failures]
        pytest.fail(
            f"Critical smoke tests failed: {names}\nDetails: {json.dumps(failures, ensure_ascii=False, indent=2)}"
        )

    assert data["passed"] is True, (
        f"Smoke overall passed flag is False:\n{json.dumps(data, ensure_ascii=False, indent=2)}"
    )
    assert r.returncode == 0, f"Smoke tool exited with rc={r.returncode}"
