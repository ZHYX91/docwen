from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "import-linter-wrapper-ci-gate-integrity-2026-07-20.md"


def _read(relative_path: str) -> str:
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_wrapper_source_and_ci_execute_the_pinned_console_callback() -> None:
    wrapper = _read("tools/run_import_linter.py")
    workflow = _read(".github/workflows/tests.yml")
    pyproject = tomllib.loads(_read("pyproject.toml"))

    assert '"-m", "importlinter.cli"' not in wrapper
    assert "from importlinter.cli import lint_imports_command" in wrapper
    assert "lint_imports_command(prog_name='lint-imports')" in wrapper
    assert '[sys.executable, "-c", _IMPORT_LINTER_BOOTSTRAP, *argv]' in wrapper
    assert "cwd=repo_root" in wrapper
    assert "env=env" in wrapper

    assert "uv sync --frozen --extra lint" in workflow
    assert "uv run python tools/run_import_linter.py" in workflow
    optional_dependencies = pyproject["project"]["optional-dependencies"]
    assert "import-linter==2.11" in optional_dependencies["lint"]
    assert "import-linter==2.11" not in optional_dependencies["test"]
