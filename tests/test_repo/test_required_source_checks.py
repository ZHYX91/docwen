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


@pytest.mark.parametrize("job", ["source-checks", "tests", "windows"])
@pytest.mark.parametrize("result", ["failure", "skipped", "cancelled", None])
def test_required_ci_job_cannot_be_skipped_or_missing(job: str, result: str | None) -> None:
    jobs = {name: {"result": "success"} for name in ("source-checks", "tests", "windows") if name != job}
    if result is not None:
        jobs[job] = {"result": result}
    with pytest.raises(ValueError, match=job):
        validate_ci_results(jobs, "workflow_dispatch")


@pytest.mark.parametrize("event", ["pull_request", "push", "workflow_dispatch"])
def test_all_events_require_the_same_source_and_platform_matrix(event: str) -> None:
    jobs = {name: {"result": "success"} for name in ("source-checks", "tests", "windows")}
    validate_ci_results(jobs, event)
    del jobs["tests"]
    with pytest.raises(ValueError, match="tests"):
        validate_ci_results(jobs, event)


def test_only_superseded_runs_for_the_same_pull_request_are_cancelled() -> None:
    workflow = yaml.load(Path(".github/workflows/tests.yml").read_text(encoding="utf-8"), Loader=yaml.BaseLoader)
    assert workflow["concurrency"] == {
        "group": "tests-${{ github.event_name }}-${{ github.event.pull_request.number || github.run_id }}",
        "cancel-in-progress": "${{ github.event_name == 'pull_request' }}",
    }


def test_ci_partitions_windows_once_and_keeps_the_platform_jobs_independent() -> None:
    workflow = yaml.load(Path(".github/workflows/tests.yml").read_text(encoding="utf-8"), Loader=yaml.BaseLoader)
    jobs = workflow["jobs"]
    roots = {"source-checks", "tests", "windows-apps", "windows-remaining"}
    assert set(jobs) == roots | {"windows", "required"}
    assert all("needs" not in jobs[name] and "if" not in jobs[name] for name in roots)
    assert jobs["tests"]["strategy"]["matrix"]["os"] == ["ubuntu-24.04", "macos-15"]
    assert jobs["tests"]["env"]["QA_SUITE"] == "platform"
    assert "QA_SHARD" not in jobs["tests"]["env"]
    assert jobs["windows"]["name"] == "Tests on windows-latest"
    assert jobs["windows"]["needs"] == ["windows-apps", "windows-remaining"]
    for job_id, group in zip(jobs["windows"]["needs"], ("apps-plugins", "remaining"), strict=True):
        job = jobs[job_id]
        assert job["name"] == f"Windows shard {group}"
        assert job["runs-on"] == "windows-latest"
        assert (job["env"]["QA_SUITE"], job["env"]["QA_SHARD"]) == ("full", group)
        assert job["steps"] == jobs["tests"]["steps"]
        assert job["outputs"] == {"artifact-id": "${{ steps.reports.outputs.artifact-id }}"}
        downloads = [
            step["with"] for step in jobs["windows"]["steps"] if "actions/download-artifact@" in step.get("uses", "")
        ]
        assert any(
            params.get("artifact-ids") == f"${{{{ needs.{job_id}.outputs.artifact-id }}}}" for params in downloads
        )
    assert jobs["required"]["if"] == "always()"
    assert jobs["required"]["needs"] == ["source-checks", "tests", "windows"]


def test_windows_scheduling_comparison_runs_only_on_explicit_dispatch() -> None:
    workflow = yaml.load(
        Path(".github/workflows/windows-shard-trial.yml").read_text(encoding="utf-8"), Loader=yaml.BaseLoader
    )
    assert set(workflow["on"]) == {"workflow_dispatch"}
