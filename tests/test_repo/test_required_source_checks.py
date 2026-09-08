from __future__ import annotations

import pytest
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
        jobs["pytest_windows_push"] = {"result": result}
    with pytest.raises(ValueError, match="pytest_windows_push"):
        validate_ci_results(jobs, "workflow_dispatch")


def test_pull_request_requires_all_platform_and_coverage_jobs() -> None:
    jobs = {
        name: {"result": "success"}
        for name in (
            "source-checks",
            "pytest_windows_pr_integration",
            "pytest_ubuntu",
            "pytest_macos",
            "coverage",
            "coverage_gui",
        )
    }
    validate_ci_results(jobs, "pull_request")
    del jobs["coverage_gui"]
    with pytest.raises(ValueError, match="coverage_gui"):
        validate_ci_results(jobs, "pull_request")
