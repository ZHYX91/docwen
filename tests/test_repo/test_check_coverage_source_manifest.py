from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _load_module():
    project_root = Path(__file__).resolve().parents[2]
    script_path = project_root / "tools" / "check_coverage_source_manifest.py"
    spec = importlib.util.spec_from_file_location("docwen_check_coverage_source_manifest", script_path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _write_repo(tmp_path: Path, *, configured: tuple[str, ...] = ("docwen_core", "docwen_runtime")) -> None:
    source_values = ", ".join(f'"{source}"' for source in configured)
    (tmp_path / "pyproject.toml").write_text(
        f"""
[tool.coverage.run]
source = [{source_values}]
""".strip(),
        encoding="utf-8",
    )
    for package_name, package_path in (
        ("docwen_core", tmp_path / "packages" / "core"),
        ("docwen_runtime", tmp_path / "packages" / "runtime"),
    ):
        (package_path / "pyproject.toml").parent.mkdir(parents=True, exist_ok=True)
        (package_path / "pyproject.toml").write_text('[project]\nname = "fixture"\n', encoding="utf-8")
        module_root = package_path / "src" / package_name
        module_root.mkdir(parents=True)
        (module_root / "__init__.py").write_text("", encoding="utf-8")


def _write_coverage_xml(tmp_path: Path, filenames: tuple[str, ...]) -> Path:
    classes = "\n".join(
        f'<class name="{Path(filename).name}" filename="{filename}"><lines /></class>' for filename in filenames
    )
    coverage_xml = tmp_path / "coverage.xml"
    coverage_xml.write_text(
        f'<coverage><packages><package name="fixture"><classes>{classes}</classes></package></packages></coverage>',
        encoding="utf-8",
    )
    return coverage_xml


def test_coverage_source_manifest_accepts_every_discovered_package(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    _write_repo(tmp_path)
    coverage_xml = _write_coverage_xml(
        tmp_path,
        (
            "packages/core/src/docwen_core/__init__.py",
            "packages/runtime/src/docwen_runtime/config.py",
        ),
    )

    exit_code = module.main([str(coverage_xml), "--repo-root", str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "reports all 2 configured packages" in captured.out


def test_coverage_source_manifest_fails_when_xml_has_no_data(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    _write_repo(tmp_path)
    coverage_xml = _write_coverage_xml(tmp_path, ())

    exit_code = module.main([str(coverage_xml), "--repo-root", str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "coverage_no_data" in captured.err


def test_coverage_source_manifest_fails_when_configured_package_is_not_reported(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    _write_repo(tmp_path)
    coverage_xml = _write_coverage_xml(tmp_path, ("packages/core/src/docwen_core/__init__.py",))

    exit_code = module.main([str(coverage_xml), "--repo-root", str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "coverage_module_not_reported: docwen_runtime" in captured.err


def test_coverage_source_manifest_fails_when_config_does_not_match_source_tree(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    module = _load_module()
    _write_repo(tmp_path, configured=("docwen_core", "docwen_phantom"))
    coverage_xml = _write_coverage_xml(tmp_path, ("packages/core/src/docwen_core/__init__.py",))

    exit_code = module.main([str(coverage_xml), "--repo-root", str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "coverage_source_manifest_missing: docwen_runtime" in captured.err
    assert "coverage_source_manifest_unknown: docwen_phantom" in captured.err
