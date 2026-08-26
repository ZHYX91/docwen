from __future__ import annotations

import argparse
import ast
import sys
from pathlib import Path

import tomlkit

EXPECTED_MARKERS = (
    "unit",
    "contract",
    "integration",
    "gui_smoke",
    "e2e",
    "pr_gate",
    "release_gate",
    "slow",
    "macos_only",
    "windows_only",
    "linux_only",
    "packaged",
    "host",
    "office",
)


def _load_pyproject(pyproject_path: Path) -> dict:
    return tomlkit.parse(pyproject_path.read_text(encoding="utf-8"))


def _discover_source_packages(repo_root: Path) -> list[str]:
    discovered: set[str] = set()
    for package_config in (repo_root / "packages").rglob("pyproject.toml"):
        source_root = package_config.parent / "src"
        if not source_root.is_dir():
            continue
        for candidate in source_root.iterdir():
            if candidate.is_dir() and (candidate / "__init__.py").is_file():
                discovered.add(candidate.name)
    return sorted(discovered)


def _load_qa_constants(qa_path: Path) -> dict[str, str]:
    tree = ast.parse(qa_path.read_text(encoding="utf-8"))
    needed = {
        "FAST_MARK_EXPR",
        "PR_GATE_MARK_EXPR",
        "RELEASE_GATE_MARK_EXPR",
        "PYTEST_BASE_ADDOPTS",
        "PYTEST_XDIST_ENV",
        "PYTEST_XDIST_WORKERS_ENV",
        "PYTEST_XDIST_DIST",
        "PYTEST_PRIMARY_MARKER_DEBT_LIMIT",
        "PYTEST_PRIMARY_MARKER_OVERLAP_LIMIT",
        "PYTEST_RUNTIME_ROOT_ENV",
        "PYTEST_REPORT_DIR_ENV",
    }
    values: dict[str, str] = {}
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if len(node.targets) != 1 or not isinstance(node.targets[0], ast.Name):
            continue
        name = node.targets[0].id
        if name not in needed:
            continue
        values[name] = ast.literal_eval(node.value)
    missing = sorted(needed - values.keys())
    if missing:
        raise ValueError(f"missing QA constants: {', '.join(missing)}")
    return values


def _contains_all(text: str, fragments: list[str], label: str, errors: list[str]) -> None:
    for fragment in fragments:
        if fragment not in text:
            errors.append(f"{label} missing expected fragment: {fragment}")


def _reject_runner_context_in_job_env(text: str, errors: list[str]) -> None:
    """Reject GitHub runner context where the workflow parser cannot resolve it."""
    in_job_env = False
    for line_number, line in enumerate(text.splitlines(), start=1):
        stripped = line.lstrip()
        if not stripped or stripped.startswith("#"):
            continue
        indentation = len(line) - len(stripped)
        if indentation == 4 and stripped == "env:":
            in_job_env = True
            continue
        if in_job_env and indentation <= 4:
            in_job_env = False
        if in_job_env and "${{ runner." in line:
            errors.append(f".github/workflows/tests.yml:{line_number} uses runner context in job-level env")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo-root", default=Path(__file__).resolve().parents[1], type=Path)
    args = parser.parse_args(argv)

    repo_root = args.repo_root.resolve()
    pyproject_path = repo_root / "pyproject.toml"
    qa_path = repo_root / "tools" / "qa.py"
    workflow_path = repo_root / ".github" / "workflows" / "tests.yml"
    docs_path = repo_root / "docs" / "testing.md"
    fixtures_readme_path = repo_root / "tests" / "fixtures" / "README.md"

    pyproject = _load_pyproject(pyproject_path)
    qa_constants = _load_qa_constants(qa_path)
    workflow_text = workflow_path.read_text(encoding="utf-8")
    docs_text = docs_path.read_text(encoding="utf-8")

    pytest_options = pyproject["tool"]["pytest"]["ini_options"]
    coverage_run_options = pyproject["tool"]["coverage"]["run"]
    coverage_options = pyproject["tool"]["coverage"]["report"]
    markers = [entry.split(":", 1)[0].strip() for entry in pytest_options["markers"]]
    coverage_sources = [str(source) for source in coverage_run_options["source"]]
    coverage_filterwarnings = [str(entry) for entry in pytest_options.get("filterwarnings", [])]
    test_dependencies = pyproject["project"]["optional-dependencies"]["test"]

    fast_expr = qa_constants["FAST_MARK_EXPR"]
    pr_gate_expr = qa_constants["PR_GATE_MARK_EXPR"]
    release_gate_expr = qa_constants["RELEASE_GATE_MARK_EXPR"]
    pytest_addopts = qa_constants["PYTEST_BASE_ADDOPTS"]
    xdist_env = qa_constants["PYTEST_XDIST_ENV"]
    xdist_workers_env = qa_constants["PYTEST_XDIST_WORKERS_ENV"]
    xdist_dist = qa_constants["PYTEST_XDIST_DIST"]
    marker_debt_limit = int(qa_constants["PYTEST_PRIMARY_MARKER_DEBT_LIMIT"])
    marker_overlap_limit = int(qa_constants["PYTEST_PRIMARY_MARKER_OVERLAP_LIMIT"])
    pytest_runtime_root_env = qa_constants["PYTEST_RUNTIME_ROOT_ENV"]
    pytest_report_dir_env = qa_constants["PYTEST_REPORT_DIR_ENV"]
    fail_under = int(coverage_options["fail_under"])

    errors: list[str] = []
    _reject_runner_context_in_job_env(workflow_text, errors)

    if f'-m "{fast_expr}"' not in pytest_options["addopts"]:
        errors.append("pyproject addopts does not match qa.py FAST_MARK_EXPR")
    if "--strict-markers" not in pytest_options["addopts"]:
        errors.append("pyproject addopts must include --strict-markers")
    if "--import-mode=importlib" not in pytest_options["addopts"]:
        errors.append("pyproject addopts must include --import-mode=importlib")
    if "--import-mode=importlib" not in pytest_addopts:
        errors.append("qa.py PYTEST_BASE_ADDOPTS must include --import-mode=importlib")
    if not any(dep.startswith("pytest-xdist") for dep in test_dependencies):
        errors.append("pyproject optional-dependencies.test missing pytest-xdist")
    if marker_debt_limit < 0:
        errors.append("qa.py primary marker debt limit must be non-negative")
    if marker_overlap_limit < 0:
        errors.append("qa.py primary marker overlap limit must be non-negative")
    if marker_overlap_limit > marker_debt_limit:
        errors.append("qa.py primary marker overlap limit cannot exceed missing marker debt limit")
    subprocess_runner_path = repo_root / "tests" / "support" / "subprocess_runner.py"
    if not subprocess_runner_path.is_file():
        errors.append("tests/support/subprocess_runner.py is missing")
    else:
        _contains_all(
            subprocess_runner_path.read_text(encoding="utf-8"),
            [
                "DEFAULT_SUBPROCESS_TIMEOUT_SECONDS = 60.0",
                "timeout: float | None = DEFAULT_SUBPROCESS_TIMEOUT_SECONDS",
                'raise ValueError("test subprocess timeout must be a positive number")',
            ],
            "tests/support/subprocess_runner.py",
            errors,
        )

    discovered_sources = _discover_source_packages(repo_root)
    if coverage_sources != sorted(set(coverage_sources)):
        errors.append("pyproject coverage source must be a unique sorted package list")
    missing_coverage_sources = sorted(set(discovered_sources) - set(coverage_sources))
    unknown_coverage_sources = sorted(set(coverage_sources) - set(discovered_sources))
    if missing_coverage_sources:
        errors.append("pyproject coverage source missing: " + ", ".join(missing_coverage_sources))
    if unknown_coverage_sources:
        errors.append("pyproject coverage source unknown: " + ", ".join(unknown_coverage_sources))
    for warning_filter in (
        "error:Module .* was never imported.*:coverage.exceptions.CoverageWarning",
        "error:No data was collected.*:coverage.exceptions.CoverageWarning",
        "error:Failed to generate report.*No data to report.*:pytest_cov.CovReportWarning",
    ):
        if warning_filter not in coverage_filterwarnings:
            errors.append(f"pyproject filterwarnings missing fail-closed coverage warning: {warning_filter}")

    for marker in EXPECTED_MARKERS:
        if marker not in markers:
            errors.append(f"pyproject markers missing: {marker}")
        if f"`{marker}`" not in docs_text:
            errors.append(f"docs/specs missing marker documentation: {marker}")

    _contains_all(
        workflow_text,
        [
            "tools/qa.py --skip-ruff --skip-pyright --suite fast",
            "tools/qa.py --skip-ruff --skip-pyright --suite pr-integration",
            "tools/qa.py --skip-ruff --skip-pyright --suite full",
            "uv run python tools/run_import_linter.py",
            f'{xdist_env}: "1"',
            f'{xdist_workers_env}: "auto"',
            pytest_runtime_root_env,
            'New-Item -ItemType Directory -Force -Path "$env:RUNNER_TEMP/docwen-pytest-runtime"',
            "libegl1",
            "PYTHONIOENCODING: utf-8",
            'PYTHONUTF8: "1"',
            "uv sync --frozen --extra test --extra dev",
            f'uv run pytest -m "{fast_expr}" -n "$env:{xdist_workers_env}" --dist {xdist_dist} --cov --cov-config=pyproject.toml --cov-report=term-missing:skip-covered',
            f"--cov-fail-under={fail_under}",
            "--basetemp",
            "cache_dir=",
            "docwen-pytest-runtime/reports/skip_report.json",
            "docwen-pytest-runtime/reports/not_collected_report.json",
            "docwen-pytest-runtime/reports/slow_report.json",
            "docwen-pytest-runtime/reports/subprocess_report.json",
            "docwen-pytest-runtime/reports/missing_marker_report.json",
            "uv run python tools/check_coverage_source_manifest.py",
            "uv run python tools/check_core_coverage.py",
            "uv run pytest packages/apps/gui/tests --cov=docwen_gui --cov-report=term-missing:skip-covered",
            "uv run python tools/check_gui_coverage.py",
            "uv run python tools/check_test_governance_consistency.py",
        ],
        ".github/workflows/tests.yml",
        errors,
    )

    _contains_all(
        docs_text,
        [
            "python tools/qa.py --skip-ruff --skip-pyright --suite pr-integration",
            "python tools/qa.py --skip-ruff --skip-pyright --suite release",
            "python tools/check_coverage_source_manifest.py coverage.xml",
            "python tools/check_core_coverage.py coverage.xml --soft-gate",
            "python tools/check_gui_coverage.py coverage-gui.xml",
            "python tools/check_test_governance_consistency.py",
            "python tools/run_import_linter.py",
            "`pytest-xdist`",
            f"`{xdist_env}`",
            f"`{xdist_workers_env}`",
            f"`{pytest_runtime_root_env}`",
            xdist_dist,
            "`skip_report.json`",
            "`not_collected_report.json`",
            "`slow_report.json`",
            "`subprocess_report.json`",
            "`missing_marker_report.json`",
            "`tests/fixtures/README.md`",
            "`tests/fixtures/golden/`",
            "`tests.support.subprocess_runner.run_subprocess`",
            "人工审 diff",
            str(fail_under),
            fast_expr,
            pr_gate_expr,
            release_gate_expr,
            pytest_addopts,
        ],
        "docs/testing.md",
        errors,
    )

    if not fixtures_readme_path.is_file():
        errors.append("tests/fixtures/README.md is missing")
    else:
        fixtures_readme_text = fixtures_readme_path.read_text(encoding="utf-8")
        _contains_all(
            fixtures_readme_text,
            [
                "`files/`",
                "`golden/`",
                "`skip_report.json`",
                "`not_collected_report.json`",
                "`slow_report.json`",
                "`subprocess_report.json`",
                "`missing_marker_report.json`",
                f"`{pytest_report_dir_env}`",
                "人工审 diff",
                "不属于 fixtures，也不属于 golden",
            ],
            "tests/fixtures/README.md",
            errors,
        )

    if errors:
        print("==> test-governance-consistency", file=sys.stderr)
        for error in errors:
            print(f"[mismatch] {error}", file=sys.stderr)
        return 1

    print("==> test-governance-consistency")
    print("ok: pyproject.toml, tools/qa.py, tests.yml, docs/testing.md and tests/fixtures/README.md are aligned")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
