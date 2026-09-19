from __future__ import annotations

from pathlib import Path

import pytest
import yaml
from tools import qa
from tools.source_checks import validate_ci_results

pytestmark = pytest.mark.unit


def test_default_qa_runs_all_required_source_checks_before_tests(monkeypatch: pytest.MonkeyPatch) -> None:
    commands: list[list[str]] = []
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _root: 0)
    monkeypatch.setattr(qa, "_run", lambda command, **_kwargs: commands.append(command) or 0)
    assert qa.main(["--skip-pytest"]) == 0
    assert any("tools/run_import_linter.py" in command for command in commands)
    assert any("tools/check_test_governance_consistency.py" in command for command in commands)
    assert any("tools/validation/check_architecture_cleanliness.py" in command for command in commands)
    assert [command[-1] for command in commands if "--pythonplatform" in command] == ["Windows", "Linux", "Darwin"]


@pytest.mark.parametrize("failed_check", ["tools/run_import_linter.py", "Linux"])
def test_required_source_failure_stops_before_pytest(monkeypatch: pytest.MonkeyPatch, failed_check: str) -> None:
    commands: list[list[str]] = []
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _root: 0)

    def run(command: list[str], **_kwargs: object) -> int:
        commands.append(command)
        return 7 if failed_check in command else 0

    monkeypatch.setattr(qa, "_run", run)
    assert qa.main(["--suite", "full"]) == 7
    assert not any("pytest" in command for command in commands)


@pytest.mark.parametrize("result", ["failure", "skipped", "cancelled", None])
def test_required_ci_job_cannot_be_skipped_or_missing(result: str | None) -> None:
    jobs = {"source-checks": {"result": "success"}}
    if result is not None:
        jobs["tests"] = {"result": result}
    with pytest.raises(ValueError, match="tests"):
        validate_ci_results(jobs, "workflow_dispatch")


@pytest.mark.parametrize("event", ["pull_request", "push", "workflow_dispatch"])
def test_all_events_require_the_same_source_and_platform_matrix(event: str) -> None:
    jobs = {name: {"result": "success"} for name in ("source-checks", "tests")}
    validate_ci_results(jobs, event)
    del jobs["tests"]
    with pytest.raises(ValueError, match="tests"):
        validate_ci_results(jobs, event)


def test_ci_has_one_full_coverage_run_and_one_run_per_other_platform() -> None:
    workflow = yaml.load(Path(".github/workflows/tests.yml").read_text(encoding="utf-8"), Loader=yaml.BaseLoader)
    jobs = workflow["jobs"]
    assert set(jobs) == {"source-checks", "tests", "required"}
    assert jobs["tests"]["needs"] == "source-checks"
    assert jobs["tests"]["strategy"]["matrix"]["include"] == [
        {"os": "windows-latest", "suite": "full", "coverage": "true"},
        {"os": "ubuntu-24.04", "suite": "platform", "coverage": "false"},
        {"os": "macos-14", "suite": "platform", "coverage": "false"},
    ]
    assert jobs["required"]["needs"] == ["source-checks", "tests"]
    assert "if" not in jobs["tests"]
