from __future__ import annotations

import json
import os
import stat
from pathlib import Path

import pytest
from tools import qa, workspace_root

pytestmark = pytest.mark.unit


def _governance_root(engineering_root: Path) -> Path:
    governed = engineering_root / ".workspace"
    governed.mkdir(parents=True)
    (governed / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    for name in workspace_root._GOVERNANCE_DIRECTORIES:
        (governed / name).mkdir()
    return governed


def test_full_qa_uses_the_explicit_pytest_runtime_root(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    commands: list[list[str]] = []
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _repo: 0)
    environments: list[dict[str, str] | None] = []

    def _record_command(command: list[str], *, env: dict[str, str] | None = None) -> int:
        commands.append(command)
        environments.append(env)
        return 0

    monkeypatch.setattr(qa, "_run", _record_command)
    workspace = _governance_root(tmp_path / "engineering")
    pytest_temp = workspace / "temp" / "docwen" / "pytest" / "isolated-pytest-temp"
    short_view = Path("W:\\")
    mounted: list[Path] = []
    unmounted: list[tuple[str, Path]] = []
    monkeypatch.setattr(qa, "_should_use_short_runtime_drive", lambda *, suite: suite == "full")
    monkeypatch.setattr(qa, "mount_short_drive", lambda root: mounted.append(root) or "W:")
    monkeypatch.setattr(qa, "drive_root", lambda _drive: short_view)
    monkeypatch.setattr(
        qa,
        "unmount_short_drive",
        lambda drive, *, expected_target: unmounted.append((drive, expected_target)),
    )

    assert (
        qa.main(
            [
                "--suite",
                "full",
                "--skip-ruff",
                "--skip-pyright",
                "--pytest-runtime-root",
                str(pytest_temp),
                "--workspace-root",
                str(workspace),
            ]
        )
        == 0
    )
    assert len(commands) == 1
    runtime_root = pytest_temp.resolve()
    assert commands[0][-4:] == [
        "--basetemp",
        str(short_view / qa.PYTEST_BASETEMP_NAME),
        "-o",
        f"cache_dir={short_view / qa.PYTEST_CACHE_NAME}",
    ]
    assert environments[0] is not None
    assert environments[0]["DOCWEN_PYTEST_RUNTIME_ROOT"] == str(short_view)
    assert environments[0]["DOCWEN_PYTEST_REPORT_DIR"] == str(short_view / qa.PYTEST_REPORT_NAME)
    assert mounted == [runtime_root]
    assert unmounted == [("W:", runtime_root)]


def test_qa_discovers_default_runtime_above_repos(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    engineering_root = tmp_path / "DocWen-Workspace"
    repo = engineering_root / "repos" / "docwen"
    repo.mkdir(parents=True)
    governed = _governance_root(engineering_root)
    monkeypatch.delenv(qa.WORKSPACE_ROOT_ENV, raising=False)

    assert qa._default_pytest_parent(repo) == governed / "temp"


def test_qa_rejects_repository_internal_pytest_runtime_root(tmp_path: Path) -> None:
    repo_root = tmp_path / "repo"
    repo_root.mkdir()
    requested_root = repo_root / ".pytest-runtime"

    with pytest.raises(ValueError, match="must be outside the repository"):
        qa._pytest_runtime_environment(repo_root.resolve(), requested_root)

    assert not requested_root.exists()


def test_qa_builds_one_workspace_contained_pytest_runtime_environment(tmp_path: Path) -> None:
    repo_root = tmp_path / "repo"
    repo_root.mkdir()
    workspace = _governance_root(tmp_path / "engineering")
    requested_root = workspace / "temp" / "docwen" / "pytest" / "caller-runtime"

    runtime_root, environment, owned = qa._pytest_runtime_environment(
        repo_root.resolve(),
        requested_root,
        workspace_root=workspace,
    )

    assert runtime_root == requested_root.resolve()
    assert environment["DOCWEN_PYTEST_RUNTIME_ROOT"] == str(runtime_root)
    assert environment["DOCWEN_PYTEST_REPORT_DIR"] == str(runtime_root / qa.PYTEST_REPORT_NAME)
    assert environment["TEMP"] == str(runtime_root / qa.PYTEST_SYSTEM_TEMP_NAME)
    assert environment["TMP"] == str(runtime_root / qa.PYTEST_SYSTEM_TEMP_NAME)
    assert environment["TMPDIR"] == str(runtime_root / qa.PYTEST_SYSTEM_TEMP_NAME)
    assert environment["DOCWEN_PYTEST_MAX_MISSING_PRIMARY_MARKERS"] == str(qa.PYTEST_PRIMARY_MARKER_DEBT_LIMIT)
    assert all(
        (runtime_root / name).is_dir()
        for name in (
            qa.PYTEST_BASETEMP_NAME,
            qa.PYTEST_CACHE_NAME,
            qa.PYTEST_REPORT_NAME,
            qa.PYTEST_SYSTEM_TEMP_NAME,
        )
    )
    assert owned is False


def test_qa_rejects_a_local_runtime_outside_the_governed_workspace(tmp_path: Path) -> None:
    repo_root = tmp_path / "repo"
    repo_root.mkdir()
    workspace = _governance_root(tmp_path / "engineering")
    external = tmp_path / "drive-root-style-runtime"

    with pytest.raises(ValueError, match="must stay inside the governed workspace temp tree"):
        qa._pytest_runtime_environment(
            repo_root.resolve(),
            external,
            workspace_root=workspace,
        )

    assert not external.exists()


def test_qa_owns_and_removes_its_default_runtime_after_success(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    governed_workspace = _governance_root(tmp_path / "engineering")
    monkeypatch.setenv(qa.WORKSPACE_ROOT_ENV, str(governed_workspace))
    monkeypatch.delenv(qa.PYTEST_RUNTIME_ROOT_ENV, raising=False)
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _repo: 0)
    observed: list[Path] = []

    def _successful_run(_command: list[str], *, env: dict[str, str] | None = None) -> int:
        assert env is not None
        observed.append(Path(env[qa.PYTEST_RUNTIME_ROOT_ENV]))
        assert (observed[-1] / qa.PYTEST_RUNTIME_LEASE).is_file()
        return 0

    monkeypatch.setattr(qa, "_run", _successful_run)
    assert qa.main(["--skip-ruff", "--skip-pyright"]) == 0
    assert len(observed) == 1
    assert not observed[0].exists()


def test_qa_retains_owned_runtime_on_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    governed_workspace = _governance_root(tmp_path / "engineering")
    monkeypatch.setenv(qa.WORKSPACE_ROOT_ENV, str(governed_workspace))
    monkeypatch.delenv(qa.PYTEST_RUNTIME_ROOT_ENV, raising=False)
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _repo: 0)
    observed: list[Path] = []

    def _failed_run(_command: list[str], *, env: dict[str, str] | None = None) -> int:
        assert env is not None
        observed.append(Path(env[qa.PYTEST_RUNTIME_ROOT_ENV]))
        return 7

    monkeypatch.setattr(qa, "_run", _failed_run)
    assert qa.main(["--skip-ruff", "--skip-pyright"]) == 7
    lease = json.loads((observed[0] / qa.PYTEST_RUNTIME_LEASE).read_text(encoding="utf-8"))
    assert lease["state"] == "retained-failure"


def test_qa_can_own_and_remove_a_new_explicit_short_root(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspace = _governance_root(tmp_path / "engineering")
    runtime_root = workspace / "temp" / "docwen" / "pytest" / "short"
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _repo: 0)
    monkeypatch.setattr(qa, "_run", lambda _command, *, env=None: 0)

    assert (
        qa.main(
            [
                "--skip-ruff",
                "--skip-pyright",
                "--pytest-runtime-root",
                str(runtime_root),
                "--own-pytest-runtime",
                "--workspace-root",
                str(workspace),
            ]
        )
        == 0
    )
    assert not runtime_root.exists()


def test_qa_marks_an_interrupted_owned_runtime_and_keeps_it_inside_workspace(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspace = _governance_root(tmp_path / "engineering")
    monkeypatch.setenv(qa.WORKSPACE_ROOT_ENV, str(workspace))
    monkeypatch.delenv(qa.PYTEST_RUNTIME_ROOT_ENV, raising=False)
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _repo: 0)
    monkeypatch.setattr(qa, "_should_use_short_runtime_drive", lambda *, suite: False)

    def _interrupt(_command: list[str], *, env: dict[str, str] | None = None) -> int:
        raise KeyboardInterrupt

    monkeypatch.setattr(qa, "_run", _interrupt)

    assert qa.main(["--skip-ruff", "--skip-pyright"]) == 130
    runtimes = list((workspace / "temp").glob(f"{qa.PYTEST_RUNTIME_PREFIX}*"))
    assert len(runtimes) == 1
    lease = json.loads((runtimes[0] / qa.PYTEST_RUNTIME_LEASE).read_text(encoding="utf-8"))
    assert lease["state"] == "retained-interrupted"


def test_qa_start_removes_an_expired_dead_owner_workspace_runtime(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspace = _governance_root(tmp_path / "engineering")
    expired = workspace / "temp" / "docwen" / "pytest" / "expired"
    expired.mkdir(parents=True)
    (expired / qa.PYTEST_RUNTIME_LEASE).write_text(
        json.dumps(
            {
                "schemaVersion": 1,
                "owner": "docwen.tools.qa",
                "kind": "pytest-runtime",
                "pid": 999_999_999,
                "createdAt": "2020-01-01T00:00:00Z",
                "state": "retained-interrupted",
                "root": str(expired.resolve()),
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv(qa.WORKSPACE_ROOT_ENV, str(workspace))
    monkeypatch.delenv(qa.PYTEST_RUNTIME_ROOT_ENV, raising=False)
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _repo: 0)
    monkeypatch.setattr(qa, "_run", lambda _command, *, env=None: 0)

    assert qa.main(["--skip-ruff", "--skip-pyright"]) == 0
    assert not expired.exists()


def test_qa_refuses_to_own_a_preexisting_explicit_root(tmp_path: Path) -> None:
    runtime_root = tmp_path / "preexisting"
    runtime_root.mkdir()

    with pytest.raises(ValueError, match="must not already exist"):
        qa._pytest_runtime_environment(
            tmp_path / "repo",
            runtime_root,
            own_requested_root=True,
        )


def test_qa_passes_its_current_interpreter_to_pyright(monkeypatch: pytest.MonkeyPatch) -> None:
    commands: list[list[str]] = []
    monkeypatch.setattr(qa, "_scan_private_symbol_usage", lambda _repo: 0)
    monkeypatch.setattr(qa, "_run", lambda command, *, env=None: commands.append(command) or 0)

    assert qa.main(["--skip-ruff", "--skip-pytest"]) == 0
    assert commands == [[qa.sys.executable, "-m", "pyright", "--level", "error", "--pythonpath", qa.sys.executable]]


def test_xdist_auto_is_bounded_but_an_explicit_worker_count_is_exact(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(qa.PYTEST_XDIST_ENV, raising=False)
    monkeypatch.setattr(qa.importlib.util, "find_spec", lambda _name: object())
    monkeypatch.delenv(qa.PYTEST_XDIST_WORKERS_ENV, raising=False)
    monkeypatch.delenv(qa.PYTEST_XDIST_MAX_PROCESSES_ENV, raising=False)

    assert qa._xdist_args() == [
        "-n",
        "auto",
        "--dist",
        "loadfile",
        "--maxprocesses",
        qa.PYTEST_XDIST_DEFAULT_MAX_PROCESSES,
    ]

    monkeypatch.setenv(qa.PYTEST_XDIST_WORKERS_ENV, "3")
    assert qa._xdist_args() == ["-n", "3", "--dist", "loadfile"]


def test_xdist_can_be_disabled_explicitly(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(qa.PYTEST_XDIST_ENV, "0")

    assert qa._xdist_args() == []


def test_xdist_auto_rejects_an_invalid_process_cap(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(qa.PYTEST_XDIST_ENV, "1")
    monkeypatch.setattr(qa.importlib.util, "find_spec", lambda _name: object())
    monkeypatch.setenv(qa.PYTEST_XDIST_WORKERS_ENV, "auto")
    monkeypatch.setenv(qa.PYTEST_XDIST_MAX_PROCESSES_ENV, "0")

    with pytest.raises(ValueError, match=qa.PYTEST_XDIST_MAX_PROCESSES_ENV):
        qa._xdist_args()


def test_qa_rejects_linked_pytest_runtime_root(tmp_path: Path) -> None:
    repo_root = tmp_path / "repo"
    repo_root.mkdir()
    target = tmp_path / "target"
    target.mkdir()
    linked_root = tmp_path / "linked-runtime"
    try:
        linked_root.symlink_to(target, target_is_directory=True)
    except OSError as error:
        pytest.skip(f"directory symlink unavailable: {error}")

    with pytest.raises(ValueError, match="must not traverse a link or reparse point"):
        qa._pytest_runtime_environment(repo_root.resolve(), linked_root)


def test_owned_runtime_cleanup_handles_a_windows_extended_length_tree(tmp_path: Path) -> None:
    runtime_root = tmp_path / "owned-runtime"
    runtime_root.mkdir()
    qa._write_runtime_lease(runtime_root, state="active")
    deep = runtime_root
    while len(str(deep)) < 300:
        deep /= "long-path-segment"
    os.makedirs(qa._windows_extended_path(deep))

    qa._cleanup_owned_runtime(runtime_root)

    assert not runtime_root.exists()


def test_owned_runtime_cleanup_clears_a_readonly_file_attribute(tmp_path: Path) -> None:
    runtime_root = tmp_path / "owned-runtime"
    runtime_root.mkdir()
    qa._write_runtime_lease(runtime_root, state="active")
    payload = runtime_root / "readonly.bin"
    payload.write_bytes(b"owned")
    payload.chmod(stat.S_IREAD)

    qa._cleanup_owned_runtime(runtime_root)

    assert not runtime_root.exists()


def test_qa_platform_mark_expressions_exclude_linux_tests_off_linux(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(qa.sys, "platform", "win32")

    assert "not linux_only" in qa._platform_fast_mark_expr()
    assert "not linux_only" in qa._platform_release_mark_expr()

    monkeypatch.setattr(qa.sys, "platform", "linux")
    assert "not linux_only" not in qa._platform_fast_mark_expr()
    assert "not linux_only" not in qa._platform_release_mark_expr()
