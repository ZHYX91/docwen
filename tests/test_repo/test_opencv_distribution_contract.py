"""Fail-closed contract for the single OpenCV distribution used by DocWen."""

from __future__ import annotations

import ast
import importlib.metadata
import re
import tomllib
from pathlib import Path

import pytest
from packaging.requirements import Requirement
from packaging.utils import canonicalize_name

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
HEADLESS_NAME = "opencv-python-headless"
HEADLESS_VERSION = "4.13.0.92"
RAPIDOCR_NAME = "rapidocr-onnxruntime"
RAPIDOCR_VERSION = "1.4.4"
UV_VERSION = "0.12.0"
MUTUALLY_EXCLUSIVE_OPENCV_DISTRIBUTIONS = {
    "opencv-python",
    "opencv-python-headless",
    "opencv-contrib-python",
    "opencv-contrib-python-headless",
}
HIGHGUI_CALLS = {
    "createButton",
    "destroyAllWindows",
    "destroyWindow",
    "displayOverlay",
    "displayStatusBar",
    "getWindowImageRect",
    "getWindowProperty",
    "imshow",
    "moveWindow",
    "namedWindow",
    "pollKey",
    "resizeWindow",
    "selectROI",
    "selectROIs",
    "setMouseCallback",
    "setWindowProperty",
    "setWindowTitle",
    "startWindowThread",
    "waitKey",
    "waitKeyEx",
}


def _requirement_name(requirement: str) -> str:
    return canonicalize_name(Requirement(requirement).name)


def _locked_packages() -> dict[str, dict[str, object]]:
    lock = tomllib.loads((ROOT / "uv.lock").read_text(encoding="utf-8"))
    packages = lock["package"]
    return {canonicalize_name(str(package["name"])): package for package in packages}


def _dependency_names(package: dict[str, object]) -> set[str]:
    dependencies = package.get("dependencies", [])
    assert isinstance(dependencies, list)
    return {
        canonicalize_name(str(dependency["name"]))
        for dependency in dependencies
        if isinstance(dependency, dict) and "name" in dependency
    }


def test_lock_selects_one_headless_opencv_distribution() -> None:
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    lock = tomllib.loads((ROOT / "uv.lock").read_text(encoding="utf-8"))
    packages = _locked_packages()

    root_requirements = {
        _requirement_name(str(requirement)): str(requirement) for requirement in project["project"]["dependencies"]
    }
    assert root_requirements[HEADLESS_NAME] == f"{HEADLESS_NAME}=={HEADLESS_VERSION}"

    selected_variants = MUTUALLY_EXCLUSIVE_OPENCV_DISTRIBUTIONS & packages.keys()
    assert selected_variants == {HEADLESS_NAME}
    assert packages[HEADLESS_NAME]["version"] == HEADLESS_VERSION
    assert packages[RAPIDOCR_NAME]["version"] == RAPIDOCR_VERSION
    assert HEADLESS_NAME in _dependency_names(packages["pdf2docx"])
    assert "opencv-python" not in _dependency_names(packages[RAPIDOCR_NAME])

    expected_exclusion = {
        "package": {"name": RAPIDOCR_NAME, "version": RAPIDOCR_VERSION},
        "dependencies": ["opencv-python"],
    }
    assert project["tool"]["uv"]["required-version"] == f"=={UV_VERSION}"
    assert project["tool"]["uv"]["exclude-dependencies"] == [expected_exclusion]
    assert lock["manifest"]["excludes"] == [expected_exclusion]


def test_installed_cv2_has_one_distribution_owner() -> None:
    import cv2

    owners = {canonicalize_name(owner) for owner in importlib.metadata.packages_distributions().get("cv2", [])}
    assert owners == {HEADLESS_NAME}
    assert importlib.metadata.version(HEADLESS_NAME) == HEADLESS_VERSION
    assert cv2.__version__ == "4.13.0"
    assert "GUI:                           NONE" in cv2.getBuildInformation()
    with pytest.raises(importlib.metadata.PackageNotFoundError):
        importlib.metadata.distribution("opencv-python")

    rapidocr_requirements = {
        str(Requirement(requirement))
        for requirement in importlib.metadata.requires(RAPIDOCR_NAME) or []
        if canonicalize_name(Requirement(requirement).name) == "opencv-python"
    }
    assert rapidocr_requirements == {"opencv-python>=4.5.1.48"}

    # Both upstream consumers must remain importable from the one shared cv2
    # implementation selected by the lock-aware environment.
    from pdf2docx import Converter
    from rapidocr_onnxruntime import RapidOCR

    assert Converter.__module__ == "pdf2docx.converter"
    assert RapidOCR.__module__ == "rapidocr_onnxruntime.main"


def test_docwen_sources_do_not_call_opencv_highgui() -> None:
    violations: list[str] = []
    for path in sorted((ROOT / "packages").glob("**/src/**/*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        cv2_aliases = {
            alias.asname or alias.name
            for node in ast.walk(tree)
            if isinstance(node, ast.Import)
            for alias in node.names
            if alias.name == "cv2"
        }
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module == "cv2":
                imported_highgui = HIGHGUI_CALLS & {alias.name for alias in node.names}
                violations.extend(f"{path.relative_to(ROOT)}:{node.lineno}:{name}" for name in sorted(imported_highgui))
            elif (
                isinstance(node, ast.Attribute)
                and isinstance(node.value, ast.Name)
                and node.value.id in cv2_aliases
                and node.attr in HIGHGUI_CALLS
            ):
                violations.append(f"{path.relative_to(ROOT)}:{node.lineno}:{node.attr}")

    assert violations == []


def test_source_and_ci_contracts_pin_frozen_uv() -> None:
    workflow_paths = (ROOT / ".github" / "workflows" / "tests.yml", ROOT / ".github" / "workflows" / "release.yml")
    for path in workflow_paths:
        source = path.read_text(encoding="utf-8")
        setup_actions = re.findall(r"uses: astral-sh/setup-uv@([0-9a-f]{40})", source)
        setup_count = len(setup_actions)
        assert setup_count > 0
        assert set(setup_actions) == {"37802adc94f370d6bfd71619e3f0bf239e1f3b78"}
        assert source.count(f'version: "{UV_VERSION}"') == setup_count
        sync_commands = re.findall(r"(?m)^\s*run:\s*(uv sync[^\r\n]*)$", source)
        assert sync_commands
        assert all(command.startswith("uv sync --frozen ") for command in sync_commands)

    public_source_docs = [
        ROOT / "README.md",
        ROOT / "DEVELOPMENT.md",
        *sorted((ROOT / "docs" / "user-guides").glob("README.*.md")),
    ]
    for path in public_source_docs:
        source = path.read_text(encoding="utf-8")
        assert UV_VERSION in source
        assert "uv sync --frozen --all-extras" in source
        assert re.search(r"(?m)^\s*pip install -e\b", source) is None
