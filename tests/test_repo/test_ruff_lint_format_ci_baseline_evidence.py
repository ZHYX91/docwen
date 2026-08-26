"""Evidence guards for VIS-2026-07-21-155 whole-repository Ruff closure."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "ruff-lint-format-ci-baseline-2026-07-21.md"


def _read(relative_path: str) -> str:
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_vis155_keeps_the_existing_lint_gate_strict_and_reachable() -> None:
    pyproject = tomllib.loads(_read("pyproject.toml"))
    workflow = _read(".github/workflows/tests.yml")
    ruff = pyproject["tool"]["ruff"]

    assert ruff["line-length"] == 120
    assert ruff["target-version"] == "py312"
    assert ruff["extend-exclude"] == ["assets", "templates", "configs", ".acceptance-runtime"]
    assert ruff["format"]["line-ending"] == "lf"
    assert ruff["lint"]["select"] == ["E", "F", "I", "W", "UP", "B", "SIM", "C4", "PIE", "RUF", "T20"]
    assert ruff["lint"]["ignore"] == ["E501", "RUF001", "RUF002", "RUF003"]

    commands = (
        "uv run ruff format --check .",
        "uv run ruff check .",
        "uv run python tools/check_test_governance_consistency.py",
        "uv run python tools/run_import_linter.py",
    )
    positions = [workflow.index(command) for command in commands]
    assert positions == sorted(positions)


def test_generated_acceptance_evidence_is_outside_the_whole_repository_ruff_gate() -> None:
    pyproject = tomllib.loads(_read("pyproject.toml"))

    assert ".acceptance-runtime" in pyproject["tool"]["ruff"]["extend-exclude"]
