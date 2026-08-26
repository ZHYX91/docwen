from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _load_module():
    project_root = Path(__file__).resolve().parents[2]
    script_path = project_root / "tools" / "check_test_governance_consistency.py"
    spec = importlib.util.spec_from_file_location("docwen_check_test_governance_consistency", script_path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _write_file(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def _build_repo(
    tmp_path: Path,
    *,
    workflow_cov_fail_under: int = 81,
    include_fixtures_readme: bool = True,
) -> None:
    _write_file(
        tmp_path / "pyproject.toml",
        """
[project]
name = "tmp"
version = "0.0.0"

[project.optional-dependencies]
test = ["pytest>=8.0.0", "pytest-cov>=4.0.0", "pytest-qt>=4.4.0", "pytest-xdist>=3.8.0"]

[tool.pytest.ini_options]
addopts = '-v --tb=short --strict-markers --import-mode=importlib -ra -m "(unit or contract) and not slow"'
filterwarnings = [
  "error:Module .* was never imported.*:coverage.exceptions.CoverageWarning",
  "error:No data was collected.*:coverage.exceptions.CoverageWarning",
  "error:Failed to generate report.*No data to report.*:pytest_cov.CovReportWarning",
]
markers = [
  "unit: u",
  "contract: c",
  "integration: i",
  "gui_smoke: g",
  "e2e: e",
  "pr_gate: p",
  "release_gate: r",
  "slow: s",
  "macos_only: m",
  "windows_only: w",
  "linux_only: l",
  "packaged: p",
  "host: h",
  "office: o",
]

[tool.coverage.run]
source = ["sample"]

[tool.coverage.report]
fail_under = 81
""".strip(),
    )
    _write_file(
        tmp_path / "tools" / "qa.py",
        """
FAST_MARK_EXPR = "(unit or contract) and not slow"
PR_GATE_MARK_EXPR = "pr_gate and (integration or gui_smoke or e2e)"
RELEASE_GATE_MARK_EXPR = "release_gate and (integration or gui_smoke or e2e)"
PYTEST_BASE_ADDOPTS = "-v --tb=short --strict-markers --import-mode=importlib -ra"
PYTEST_XDIST_ENV = "DOCWEN_PYTEST_XDIST"
PYTEST_XDIST_WORKERS_ENV = "DOCWEN_PYTEST_XDIST_WORKERS"
PYTEST_XDIST_DIST = "loadfile"
PYTEST_PRIMARY_MARKER_DEBT_LIMIT = 872
PYTEST_PRIMARY_MARKER_OVERLAP_LIMIT = 82
PYTEST_RUNTIME_ROOT_ENV = "DOCWEN_PYTEST_RUNTIME_ROOT"
PYTEST_REPORT_DIR_ENV = "DOCWEN_PYTEST_REPORT_DIR"
""".strip(),
    )
    _write_file(
        tmp_path / "tests" / "support" / "subprocess_runner.py",
        """
DEFAULT_SUBPROCESS_TIMEOUT_SECONDS = 60.0

def run_subprocess(timeout: float | None = DEFAULT_SUBPROCESS_TIMEOUT_SECONDS):
    if timeout is None or timeout <= 0:
        raise ValueError("test subprocess timeout must be a positive number")
""".strip(),
    )
    _write_file(
        tmp_path / ".github" / "workflows" / "tests.yml",
        f"""
uv run python tools/qa.py --skip-ruff --skip-pyright --suite fast
uv run python tools/qa.py --skip-ruff --skip-pyright --suite pr-integration
uv run python tools/qa.py --skip-ruff --skip-pyright --suite full
uv run python tools/run_import_linter.py
DOCWEN_PYTEST_XDIST: "1"
DOCWEN_PYTEST_XDIST_WORKERS: "auto"
DOCWEN_PYTEST_RUNTIME_ROOT
New-Item -ItemType Directory -Force -Path "$env:RUNNER_TEMP/docwen-pytest-runtime"
uv run pytest -m "(unit or contract) and not slow" -n "$env:DOCWEN_PYTEST_XDIST_WORKERS" --dist loadfile --cov --cov-config=pyproject.toml --cov-report=term-missing:skip-covered --cov-report="xml:$env:RUNNER_TEMP/docwen-pytest-runtime/coverage.xml" --cov-report="html:$env:RUNNER_TEMP/docwen-pytest-runtime/htmlcov" --cov-fail-under={workflow_cov_fail_under} --basetemp "$env:RUNNER_TEMP/docwen-pytest-runtime/basetemp" -o "cache_dir=$env:RUNNER_TEMP/docwen-pytest-runtime/cache"
docwen-pytest-runtime/reports/skip_report.json
docwen-pytest-runtime/reports/not_collected_report.json
docwen-pytest-runtime/reports/slow_report.json
docwen-pytest-runtime/reports/subprocess_report.json
docwen-pytest-runtime/reports/missing_marker_report.json
uv run python tools/check_coverage_source_manifest.py "$env:RUNNER_TEMP/docwen-pytest-runtime/coverage.xml"
uv run python tools/check_core_coverage.py "$env:RUNNER_TEMP/docwen-pytest-runtime/coverage.xml" --soft-gate
uv run pytest packages/apps/gui/tests --cov=docwen_gui --cov-report=term-missing:skip-covered --cov-report="xml:$env:RUNNER_TEMP/docwen-pytest-runtime/coverage-gui.xml" --cov-report="html:$env:RUNNER_TEMP/docwen-pytest-runtime/htmlcov-gui" -o addopts="-v --tb=short --strict-markers --import-mode=importlib -ra" --basetemp "$env:RUNNER_TEMP/docwen-pytest-runtime/basetemp" -o "cache_dir=$env:RUNNER_TEMP/docwen-pytest-runtime/cache"
uv run python tools/check_gui_coverage.py "$env:RUNNER_TEMP/docwen-pytest-runtime/coverage-gui.xml"
uv run python tools/check_gui_coverage.py coverage-gui.xml
uv run python tools/check_test_governance_consistency.py
libegl1
PYTHONIOENCODING: utf-8
PYTHONUTF8: "1"
uv sync --frozen --extra test --extra dev
""".strip(),
    )
    _write_file(
        tmp_path / "docs" / "testing.md",
        """
`unit` `contract` `integration` `gui_smoke` `e2e` `pr_gate` `release_gate` `slow` `macos_only` `windows_only` `linux_only` `packaged` `host` `office`
`python tools/qa.py --skip-ruff --skip-pyright --suite pr-integration`
`python tools/qa.py --skip-ruff --skip-pyright --suite release`
`pytest -m "(unit or contract) and not slow" --cov --cov-config=pyproject.toml --cov-report=term-missing:skip-covered --cov-report=xml --cov-report=html --cov-fail-under=81`
`python tools/check_coverage_source_manifest.py coverage.xml`
`python tools/check_core_coverage.py coverage.xml --soft-gate`
`pytest packages/apps/gui/tests --cov=docwen_gui --cov-report=term-missing:skip-covered --cov-report=xml:coverage-gui.xml --cov-report=html:htmlcov-gui -o addopts="-v --tb=short --strict-markers --import-mode=importlib -ra"`
`python tools/check_gui_coverage.py coverage-gui.xml`
`python tools/check_test_governance_consistency.py`
`python tools/run_import_linter.py`
`pytest-xdist`
`DOCWEN_PYTEST_XDIST`
`DOCWEN_PYTEST_XDIST_WORKERS`
`DOCWEN_PYTEST_RUNTIME_ROOT`
loadfile
`skip_report.json`
`not_collected_report.json`
`slow_report.json`
`subprocess_report.json`
`missing_marker_report.json`
`tests/fixtures/README.md`
`tests/fixtures/golden/`
`tests.support.subprocess_runner.run_subprocess`
人工审 diff
-v --tb=short --strict-markers --import-mode=importlib -ra
(unit or contract) and not slow
pr_gate and (integration or gui_smoke or e2e)
release_gate and (integration or gui_smoke or e2e)
81
""".strip(),
    )
    if include_fixtures_readme:
        _write_file(
            tmp_path / "tests" / "fixtures" / "README.md",
            """
`files/`
`golden/`
`skip_report.json`
`not_collected_report.json`
`slow_report.json`
`subprocess_report.json`
`missing_marker_report.json`
`DOCWEN_PYTEST_REPORT_DIR`
人工审 diff
不属于 fixtures，也不属于 golden
""".strip(),
        )
    _write_file(
        tmp_path / "packages" / "sample" / "pyproject.toml",
        '[project]\nname = "sample"\nversion = "0.0.0"\n',
    )
    _write_file(tmp_path / "packages" / "sample" / "src" / "sample" / "__init__.py", "")


def test_check_test_governance_consistency_passes_for_current_repo(
    capsys: pytest.CaptureFixture[str],
) -> None:
    module = _load_module()
    project_root = Path(__file__).resolve().parents[2]

    exit_code = module.main(["--repo-root", str(project_root)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert (
        "ok: pyproject.toml, tools/qa.py, tests.yml, docs/testing.md and tests/fixtures/README.md are aligned"
        in captured.out
    )


def test_check_test_governance_consistency_reports_threshold_drift(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    _build_repo(tmp_path, workflow_cov_fail_under=80)

    exit_code = module.main(["--repo-root", str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "--cov-fail-under=81" in captured.err


def test_check_test_governance_consistency_reports_qa_import_mode_drift(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    _build_repo(tmp_path)
    qa_path = tmp_path / "tools" / "qa.py"
    qa_path.write_text(
        qa_path.read_text(encoding="utf-8").replace(" --import-mode=importlib", ""),
        encoding="utf-8",
    )

    exit_code = module.main(["--repo-root", str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "qa.py PYTEST_BASE_ADDOPTS must include --import-mode=importlib" in captured.err


def test_check_test_governance_consistency_reports_missing_fixtures_readme(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    _build_repo(tmp_path, include_fixtures_readme=False)

    exit_code = module.main(["--repo-root", str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "tests/fixtures/README.md is missing" in captured.err


def test_check_test_governance_consistency_rejects_runner_context_in_job_env(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    _build_repo(tmp_path)
    workflow_path = tmp_path / ".github" / "workflows" / "tests.yml"
    workflow_path.write_text(
        workflow_path.read_text(encoding="utf-8")
        + "\n    env:\n      INVALID_RUNTIME_ROOT: ${{ runner.temp }}/invalid\n",
        encoding="utf-8",
    )

    exit_code = module.main(["--repo-root", str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "uses runner context in job-level env" in captured.err
