from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest
from tests.support.subprocess_runner import DEFAULT_SUBPROCESS_TIMEOUT_SECONDS, run_subprocess

pytestmark = pytest.mark.unit


def test_test_subprocess_runner_applies_a_bounded_default(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    report_dir = tmp_path / "reports"
    monkeypatch.setenv("DOCWEN_PYTEST_REPORT_DIR", str(report_dir))

    completed = run_subprocess([sys.executable, "-c", "print('bounded')"], cwd=tmp_path)

    assert completed.returncode == 0
    assert completed.stdout.strip() == "bounded"
    records = [
        json.loads(line)
        for path in report_dir.glob("subprocess_runs-*.jsonl")
        for line in path.read_text(encoding="utf-8").splitlines()
    ]
    assert len(records) == 1
    assert records[0]["timeout_seconds"] == DEFAULT_SUBPROCESS_TIMEOUT_SECONDS
    assert records[0]["timed_out"] is False


@pytest.mark.parametrize("timeout", [None, 0, -1])
def test_test_subprocess_runner_rejects_disabled_timeout(timeout: float | None) -> None:
    with pytest.raises(ValueError, match="timeout must be a positive number"):
        run_subprocess([sys.executable, "-c", "pass"], timeout=timeout)


def test_test_subprocess_runner_records_timeout(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    report_dir = tmp_path / "reports"
    monkeypatch.setenv("DOCWEN_PYTEST_REPORT_DIR", str(report_dir))

    with pytest.raises(subprocess.TimeoutExpired):
        run_subprocess(
            [sys.executable, "-c", "import time; time.sleep(30)"],
            cwd=tmp_path,
            timeout=0.1,
        )

    record_path = next(report_dir.glob("subprocess_runs-*.jsonl"))
    record = json.loads(record_path.read_text(encoding="utf-8").strip())
    assert record["returncode"] is None
    assert record["timed_out"] is True
    assert record["timeout_seconds"] == 0.1
