"""Build and workflow orchestration contracts around packaged verifiers."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
import yaml
from tests.support.release_packaging import (
    use_compact_pymupdf_layout_manifest,
    write_packaged_common_resources,
    write_packaged_gui_assets,
)
from tests.support.subprocess_runner import run_subprocess

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _use_compact_manifest(monkeypatch: pytest.MonkeyPatch) -> None:
    use_compact_pymupdf_layout_manifest(monkeypatch)


class _FakeBuildLogger:
    def __init__(self) -> None:
        self.errors: list[str] = []
        self.infos: list[str] = []

    def debug(self, _message: str) -> None:
        pass

    def error(self, message: str) -> None:
        self.errors.append(message)

    def end_step(self) -> None:
        pass

    def info(self, message: str) -> None:
        self.infos.append(message)

    def print_summary(self) -> None:
        pass

    def start_step(self, _message: str) -> None:
        pass

    def warning(self, _message: str) -> None:
        pass


def _write_build_required_docs(deploy_dir: Path) -> None:
    (deploy_dir / "README.md").write_text("readme", encoding="utf-8")
    (deploy_dir / "LICENSE").write_text("license", encoding="utf-8")


@pytest.mark.slow
def test_packaged_verifiers_remain_directly_executable() -> None:
    for script_path, expected_options in (
        ("scripts/release/verify_packaged_cli.py", ("--binary-dir", "--ocr-smoke")),
        ("scripts/release/verify_packaged_gui.py", ("--binary-dir",)),
    ):
        proc = run_subprocess(
            [sys.executable, script_path, "--help"],
            timeout=60,
        )
        assert proc.returncode == 0, proc.stderr
        assert all(option in proc.stdout for option in expected_options)


def test_build_verify_uses_packaged_resource_preflight(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.build import build

    deploy_dir = tmp_path / "deploy"
    deploy_dir.mkdir()
    write_packaged_common_resources(deploy_dir)
    _write_build_required_docs(deploy_dir)
    fake_logger = _FakeBuildLogger()
    monkeypatch.setattr(build, "logger", fake_logger)

    assert build.verify_build(deploy_dir, with_cli=False, with_gui=False) is True


def test_build_verify_fails_when_packaged_config_file_is_missing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from scripts.build import build

    deploy_dir = tmp_path / "deploy"
    deploy_dir.mkdir()
    write_packaged_common_resources(deploy_dir)
    _write_build_required_docs(deploy_dir)
    (deploy_dir / "configs" / "numbering" / "add.toml").unlink()
    fake_logger = _FakeBuildLogger()
    monkeypatch.setattr(build, "logger", fake_logger)

    assert build.verify_build(deploy_dir, with_cli=False, with_gui=False) is False
    assert any("build_configs_missing" in message for message in fake_logger.errors)


def test_build_verify_forces_packaged_settings_archive_and_frozen_smoke(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.build import build
    from scripts.release import verify_packaged_gui

    deploy_dir = tmp_path / "dist"
    deploy_dir.mkdir()
    write_packaged_common_resources(deploy_dir)
    write_packaged_gui_assets(deploy_dir)
    _write_build_required_docs(deploy_dir)
    (deploy_dir / build.EXE_NAME).write_bytes(b"placeholder")
    calls: list[list[str]] = []
    monkeypatch.setattr(verify_packaged_gui, "main", lambda args: calls.append(list(args)) or 0)
    monkeypatch.setattr(build, "logger", _FakeBuildLogger())

    assert build.verify_build(deploy_dir, with_cli=False, with_gui=True) is True
    assert calls == [
        [
            "--binary-dir",
            str(deploy_dir),
            "--binary-name",
            build.EXE_NAME,
            "--settings-smoke",
        ]
    ]


def _release_workflow() -> dict[str, object]:
    value = yaml.load(
        Path(".github/workflows/release.yml").read_text(encoding="utf-8"),
        Loader=yaml.BaseLoader,
    )
    assert isinstance(value, dict)
    return value


def _commands(job: dict[str, object]) -> str:
    steps = job["steps"]
    assert isinstance(steps, list)
    return "\n".join(str(step.get("run", "")) for step in steps if isinstance(step, dict))


def test_release_workflow_builds_each_supported_package_twice_and_runs_packaged_gates() -> None:
    workflow = _release_workflow()
    jobs = workflow["jobs"]
    assert isinstance(jobs, dict)
    windows = jobs["build-windows"]
    linux = jobs["build-linux"]
    assert isinstance(windows, dict)
    assert isinstance(linux, dict)

    for job in (windows, linux):
        strategy = job["strategy"]
        assert isinstance(strategy, dict)
        matrix = strategy["matrix"]
        assert isinstance(matrix, dict)
        assert matrix["replica"] == ["a", "b"]

    windows_commands = _commands(windows)
    assert "scripts/release/build_production_candidate.py" in windows_commands
    assert "--proofread-report-smoke" in windows_commands
    assert "verify_packaged_gui.py" in windows_commands
    assert "--settings-smoke" in windows_commands
    assert windows_commands.count("if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }") == 3
    assert '$env:PYTHONUTF8 = "1"' in windows_commands
    assert '$env:PYTHONIOENCODING = "utf-8"' in windows_commands

    linux_commands = _commands(linux)
    assert linux_commands.count("scripts/release/linux_archive.py") == 2
    assert "--proofread-report-smoke" in linux_commands
    assert "verify_packaged_gui.py" in linux_commands
    assert "--settings-smoke" in linux_commands
    linux_dependency_step = next(
        step for step in linux["steps"] if step.get("name") == "Install Linux GUI dependencies"
    )
    assert "libegl1" in linux_dependency_step["run"]

    for job_name, job in jobs.items():
        assert isinstance(job, dict), job_name
        environment = job.get("env", {})
        assert isinstance(environment, dict), job_name
        assert all("${{ runner." not in str(value) for value in environment.values()), job_name


def test_release_workflow_has_a_read_only_preflight_and_fixed_immutable_publication_boundary() -> None:
    workflow_document = Path(".github/workflows/release.yml").read_text(encoding="utf-8")
    workflow = _release_workflow()
    jobs = workflow["jobs"]
    assert isinstance(jobs, dict)
    verify = jobs["verify-release"]
    publish = jobs["publish"]
    assert isinstance(verify, dict)
    assert isinstance(publish, dict)

    assert workflow["concurrency"] == {
        "group": "release-${{ github.repository }}",
        "cancel-in-progress": "false",
    }
    assert publish["if"] == "github.event_name == 'push'"
    assert "actions: read" in workflow_document
    assert "attestations: write" in workflow_document
    assert "id-token: write" in workflow_document
    assert "contents: write" in workflow_document
    assert "DOCWEN_PYTEST_RUNTIME_ROOT: ${{ runner.temp }}/docwen-pytest-runtime" in workflow_document
    assert "gh release create" in _commands(publish)
    assert "gh release upload" not in workflow_document
    assert "--clobber" not in workflow_document
    assert "isImmutable" not in _commands(publish)
    assert ".immutable == true" in _commands(publish)
    assert "/releases/tags/$RELEASE_VERSION" in _commands(publish)
    assert 'case "$release_status" in' in _commands(publish)
    assert "404)" in _commands(publish)
    assert "git ls-remote --exit-code" in _commands(publish)
    steps = publish["steps"]
    assert isinstance(steps, list)
    release_state = next(step for step in steps if isinstance(step, dict) and step.get("id") == "release_state")
    assert 'if gh release view "$RELEASE_VERSION"' not in str(release_state.get("run", ""))
    assert '((.published_at | type) == "string")' in str(release_state.get("run", ""))
    assert "((.published_at | length) > 0)" in str(release_state.get("run", ""))
    assert "gh attestation verify" in _commands(publish)
    assert "cmp " in _commands(verify)
    assert "artifact-ids: ${{ needs.verify-release.outputs.artifact_id }}" in workflow_document
    assert "actions/checkout@" not in workflow_document.split("\n  publish:\n", 1)[1]

    for line in workflow_document.splitlines():
        if "uses:" in line:
            action = line.split("uses:", 1)[1].split("#", 1)[0].strip()
            assert action.rsplit("@", 1)[1].isalnum()
            assert len(action.rsplit("@", 1)[1]) == 40


def test_release_workflow_publishes_supported_windows_and_ubuntu_assets() -> None:
    workflow = Path(".github/workflows/release.yml").read_text(encoding="utf-8")
    packaging = Path("docs/packaging.md").read_text(encoding="utf-8")
    readme = Path("README.md").read_text(encoding="utf-8")
    project_metadata = Path("pyproject.toml").read_text(encoding="utf-8")
    assert "DocWenCLI-windows-x64.zip" not in workflow
    assert "DocWen-windows-x64.zip" in workflow
    assert "DocWenCLI-${RELEASE_VERSION}-linux-x64.tar.gz" in workflow
    assert "DocWen-${RELEASE_VERSION}-linux-x64.tar.gz" in workflow
    assert "DocWen-macos" not in workflow
    assert "DocWen 0.9 publishes one Windows x64 package and two Ubuntu 24.04 x64 packages" in packaging
    assert "DocWen 0.9 正式发布" in packaging
    assert (
        "The [0.9.0 Release](https://github.com/ZHYX91/docwen/releases/tag/0.9.0) publishes one Windows x64 GUI+CLI package"
        in readme
    )
    assert "No 0.9 Release is published yet" not in readme
    assert "Ubuntu 24.04 x64 GUI+CLI" in readme
    assert '"Operating System :: Microsoft :: Windows"' in project_metadata
    assert '"Operating System :: POSIX :: Linux"' in project_metadata
    assert '"Operating System :: MacOS :: MacOS X"' not in project_metadata


def test_build_main_exits_when_verification_fails(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.build import build

    deploy_dir = tmp_path / "deploy"
    deploy_dir.mkdir()
    fake_logger = _FakeBuildLogger()

    def fake_init_logger() -> _FakeBuildLogger:
        build.logger = fake_logger
        return fake_logger

    monkeypatch.setattr(build, "init_logger", fake_init_logger)
    monkeypatch.setattr(build, "build_app", lambda **_kwargs: ("0.0.0", deploy_dir))
    monkeypatch.setattr(build, "verify_build", lambda _deploy_dir, **_kwargs: False)
    monkeypatch.setattr(sys, "argv", ["build.py", "--cli-only"])

    with pytest.raises(SystemExit) as exc_info:
        build.main()

    assert exc_info.value.code == 1
    assert "构建验证失败!" in fake_logger.errors
    assert not any("构建成功完成" in message for message in fake_logger.infos)
