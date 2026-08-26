"""Repository guards for source-tree console entry dependency wiring."""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tomllib
import zipfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _project_dependencies(pyproject_path: Path) -> set[str]:
    data = tomllib.loads(pyproject_path.read_text(encoding="utf-8"))
    dependencies = data["project"].get("dependencies", [])
    names: set[str] = set()
    for dependency in dependencies:
        requirement = dependency.split(";", 1)[0].split("[", 1)[0].strip().lower()
        names.add(re.split(r"\s*(?:===|==|~=|!=|<=|>=|<|>)", requirement, maxsplit=1)[0].strip())
    return names


def _package_name_for_import(import_name: str) -> str:
    return import_name.replace("_", "-")


def test_root_console_scripts_depend_on_bundle_composition_root() -> None:
    """Root `docwen`/`docwen-gui` scripts must install the bundle graph."""
    pyproject = PROJECT_ROOT / "pyproject.toml"
    data = tomllib.loads(pyproject.read_text(encoding="utf-8"))

    assert data["project"]["scripts"]["docwen"] == "docwen_bundle.cli_entry:main"
    assert data["project"]["scripts"]["docwen-gui"] == "docwen_bundle.gui_entry:main"
    assert "docwen-bundle" in _project_dependencies(pyproject)


def test_root_distribution_builds_without_flat_layout_discovery(tmp_path: Path) -> None:
    """The root metapackage must build beside workspace asset directories."""
    probe = tmp_path / "root-distribution"
    probe.mkdir()
    shutil.copy2(PROJECT_ROOT / "pyproject.toml", probe / "pyproject.toml")
    shutil.copy2(PROJECT_ROOT / "README.md", probe / "README.md")

    # These names reproduce the setuptools flat-layout failure from the real
    # checkout. None belongs to the root distribution; workspace members own
    # all importable runtime packages.
    for name in ("assets", "configs", "i18n", "models", "packages", "samples", "templates"):
        directory = probe / name
        directory.mkdir()
        (directory / "__init__.py").write_text("", encoding="utf-8")

    wheel_dir = tmp_path / "wheelhouse"
    wheel_dir.mkdir()
    command = [
        sys.executable,
        "-c",
        (
            "import os, pathlib, setuptools.build_meta, sys; "
            "os.chdir(sys.argv[1]); "
            "print(setuptools.build_meta.build_wheel(sys.argv[2]))"
        ),
        str(probe),
        str(wheel_dir),
    ]
    completed = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        timeout=120,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
    wheels = list(wheel_dir.glob("docwen-0.9.0-*.whl"))
    assert len(wheels) == 1
    with zipfile.ZipFile(wheels[0]) as archive:
        names = archive.namelist()
        assert not any(name.startswith("assets/") for name in names)
        assert not any(name.startswith("packages/") for name in names)
        entry_points = archive.read("docwen-0.9.0.dist-info/entry_points.txt").decode("utf-8")
    assert "docwen = docwen_bundle.cli_entry:main" in entry_points
    assert "docwen-gui = docwen_bundle.gui_entry:main" in entry_points


def test_bundle_dependencies_cover_default_runtime_plugins() -> None:
    """The bundle package must not rely on accidental editable paths for plugins."""
    from docwen_bundle.runtime_factory import _DEFAULT_PLUGIN_IMPORTS

    bundle_dependencies = _project_dependencies(PROJECT_ROOT / "packages" / "bundle" / "pyproject.toml")
    missing = sorted(
        _package_name_for_import(import_name)
        for import_name in _DEFAULT_PLUGIN_IMPORTS
        if _package_name_for_import(import_name) not in bundle_dependencies
    )

    assert missing == []


def test_source_entrypoint_release_gate_is_recorded_in_current_specs() -> None:
    release_baseline = (PROJECT_ROOT / "docs" / "packaging.md").read_text(encoding="utf-8")
    gate_spec = (PROJECT_ROOT / "docs" / "testing.md").read_text(encoding="utf-8")

    assert "release_gate and (integration or gui_smoke or e2e)" in gate_spec
    assert "tests/e2e/test_source_tree_entrypoints.py" in release_baseline
    assert "docwen --help" in release_baseline
    assert "docwen-gui" in release_baseline
    assert "source-tree smoke" in release_baseline


def test_source_entrypoint_smoke_requires_installed_console_scripts() -> None:
    """Release-gate entrypoint smoke must not silently fall back to direct imports."""
    source = (PROJECT_ROOT / "tests" / "e2e" / "test_source_tree_entrypoints.py").read_text(encoding="utf-8")

    assert "Path(sys.executable).parent / script_name" in source
    assert "pytest.fail" in source
    assert "must execute the installed console script" in source
    assert "direct Python import fallback" in source
    assert "docwen_bundle.{module}" not in source
    assert "from docwen_bundle." not in source
