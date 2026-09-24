"""Real CLI envelope and publication boundaries for spreadsheet diagnostics."""

import json
import os
import sys
from pathlib import Path

import pytest
from tests.support.cli import bundle_cli_command
from tests.support.formula_cache_fixture import write_formula_cache_fixture
from tests.support.subprocess_runner import run_subprocess

pytestmark = [
    pytest.mark.e2e,
    pytest.mark.pr_gate,
    pytest.mark.release_gate,
    pytest.mark.skipif(sys.platform == "darwin", reason="Primary conversion unavailable on macOS"),
]


def _run(source: Path, output: Path, target: str, *, json_mode: bool = True):
    environment = os.environ.copy()
    environment["DOCWEN_DATA_DIR"] = str(source.parent / "profile")
    arguments = [*bundle_cli_command(), "convert", str(source), "--to", target, "--output-dir", str(output)]
    if json_mode:
        arguments.append("--json")
    return run_subprocess(
        arguments,
        cwd=Path(__file__).resolve().parents[2],
        env=environment,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )


@pytest.mark.parametrize("suffix", ["csv", "tsv"])
def test_rejected_long_text_does_not_publish_partial_result(tmp_path: Path, suffix: str) -> None:
    source = tmp_path / f"long.{suffix}"
    source.write_text("x" * 40000, encoding="utf-8")
    output = tmp_path / "published"
    process = _run(source, output, "xlsx")
    payload = json.loads(process.stdout)
    assert process.returncode != 0 and payload["success"] is False
    assert "CELL-TEXT-TOO-LONG" in json.dumps(payload)
    assert "40000" in json.dumps(payload)
    assert not output.exists() or not list(output.iterdir())


@pytest.mark.parametrize("json_mode", [True, False])
def test_formula_cache_warning_survives_cli_result(tmp_path: Path, json_mode: bool) -> None:
    source = tmp_path / "cache.xlsx"
    write_formula_cache_fixture(source)
    process = _run(source, tmp_path / "published", "csv", json_mode=json_mode)
    assert process.returncode == 0, process.stderr or process.stdout
    if json_mode:
        payload = json.loads(process.stdout)
        assert payload["success"] is True
        assert any("FORMULA-CACHE-UNAVAILABLE" in warning["code"] for warning in payload["warnings"])
        text = json.dumps(payload["warnings"])
    else:
        text = process.stderr
    assert "Calc!B1" in text and "Calc!A1" not in text
