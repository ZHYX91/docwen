"""QA's selected workspace survives fresh interpreters and linked worktrees."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from tools import qa, workspace_root

pytestmark = pytest.mark.contract

_SOURCE_ROOT = Path(__file__).resolve().parents[2]
_CHILD = """
import json, os, sys
from pathlib import Path
from tools.workspace_root import WORKSPACE_ROOT_ENV, resolve_workspace_root
from scripts.release import build_msix

repo, work, outside = map(Path, sys.argv[1:])
build_msix._REPO_ROOT = repo
selected = resolve_workspace_root(repo)
checked = build_msix._safe_work_path(work)
try:
    build_msix._safe_work_path(outside)
except Exception as error:
    refusal = str(error)
else:
    raise AssertionError("MSIX must still reject work outside the selected temp tree")
print(json.dumps({
    "configured": os.environ[WORKSPACE_ROOT_ENV],
    "selected": str(selected), "checked": str(checked), "refusal": refusal,
}))
"""


def _repositories(tmp_path: Path, topology: str) -> tuple[Path, Path, Path]:
    repo = tmp_path / "a" / "repos" / "docwen"
    repo.mkdir(parents=True)

    def git(*args: str) -> None:
        subprocess.run(["git", "-C", str(repo), *args], check=True, capture_output=True, timeout=15)

    git("init")
    git("-c", "user.name=Test", "-c", "user.email=test@example.invalid", "commit", "--allow-empty", "-m", "initial")
    discovered = workspace_root.initialize_workspace(tmp_path / "a" / "runtime", [repo])
    other = tmp_path / "b"
    other.mkdir()
    if topology == "worktree":
        target = other / "worktree"
        git("worktree", "add", "--detach", str(target))
        explicit = workspace_root.initialize_workspace(other / "runtime", [target])
    else:
        target = repo
        explicit = workspace_root.initialize_workspace(other / "runtime", [])
    return target, discovered, explicit


@pytest.mark.parametrize("topology", ["checkout", "worktree"])
@pytest.mark.parametrize("selection", ["explicit", "environment", "same", "conflict", "discovered"])
def test_selected_workspace_reaches_fresh_msix_process(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, topology: str, selection: str
) -> None:
    repo, discovered, explicit = _repositories(tmp_path, topology)
    monkeypatch.delenv(qa.PYTEST_RUNTIME_ROOT_ENV, raising=False)
    monkeypatch.delenv(qa.WORKSPACE_ROOT_ENV, raising=False)
    if selection in {"environment", "same"}:
        monkeypatch.setenv(qa.WORKSPACE_ROOT_ENV, str(explicit))
    elif selection == "conflict":
        monkeypatch.setenv(qa.WORKSPACE_ROOT_ENV, str(discovered))
    inherited = os.environ.get(qa.WORKSPACE_ROOT_ENV)
    source_path = os.pathsep.join(filter(None, [str(_SOURCE_ROOT), os.environ.get("PYTHONPATH")]))
    monkeypatch.setenv("PYTHONPATH", source_path)
    requested = explicit if selection in {"explicit", "same", "conflict"} else None
    expected = discovered if selection == "discovered" else explicit

    runtime, environment, _owned = qa._pytest_runtime_environment(repo, None, workspace_root=requested)
    work = runtime / "build" / "msix"
    completed = subprocess.run(
        [sys.executable, "-c", _CHILD, str(repo), str(work), str(expected / "build" / "forbidden")],
        cwd=repo,
        env=environment,
        capture_output=True,
        text=True,
        timeout=20,
        check=True,
    )

    observed = json.loads(completed.stdout)
    assert observed["configured"] == observed["selected"] == str(expected)
    assert observed["checked"] == str(work)
    assert "msix_work_root_outside_workspace_temp" in observed["refusal"]
    assert runtime.is_relative_to(expected / "temp")
    assert os.environ.get(qa.WORKSPACE_ROOT_ENV) == inherited


def test_explicit_workspace_is_also_used_by_source_checks(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    repo, inherited, selected = _repositories(tmp_path, "worktree")
    monkeypatch.setenv(qa.WORKSPACE_ROOT_ENV, str(inherited))
    monkeypatch.setattr(qa, "__file__", str(repo / "tools" / "qa.py"))
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _repo: 0)
    monkeypatch.setattr(qa.source_checks, "architecture_steps", lambda: [("probe", ["probe"], True)])
    environments = []

    def run(_command: list[str], *, env: dict[str, str] | None = None) -> int:
        assert env is not None
        environments.append(env)
        return 0

    monkeypatch.setattr(qa, "_run", run)
    result = qa.main(["--workspace-root", str(selected), "--skip-ruff", "--skip-pyright", "--skip-pytest"])

    assert result == 0
    assert len(environments) == 1
    assert environments[0][qa.WORKSPACE_ROOT_ENV] == str(selected)
    assert os.environ[qa.WORKSPACE_ROOT_ENV] == str(inherited)
