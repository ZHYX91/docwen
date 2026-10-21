"""Keep image decoders and development isolation above their security floors."""

from __future__ import annotations

import importlib.metadata
import tomllib
from pathlib import Path

import pytest
from packaging.requirements import Requirement
from packaging.version import Version

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]


@pytest.mark.parametrize("owner", [".", "packages/core", "packages/plugins/image"])
def test_owned_pillow_requirements_exclude_affected_decoders(owner: str) -> None:
    project = tomllib.loads((ROOT / owner / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    requirements = [Requirement(value) for value in project["dependencies"]]
    pillow = next(requirement for requirement in requirements if requirement.name.lower() == "pillow")

    assert "12.3.0" not in pillow.specifier
    assert "12.9.9" not in pillow.specifier
    assert "13.0.0" in pillow.specifier
    assert "13.0.1" in pillow.specifier
    assert "14.0.0" not in pillow.specifier


@pytest.mark.parametrize("dependency,floor", [("pillow", "13.0.0"), ("virtualenv", "21.14.9")])
def test_locked_and_installed_dependencies_meet_security_floor(dependency: str, floor: str) -> None:
    lock = tomllib.loads((ROOT / "uv.lock").read_text(encoding="utf-8"))
    selected = next(package["version"] for package in lock["package"] if package["name"] == dependency)

    assert Version(selected) >= Version(floor)
    assert importlib.metadata.version(dependency) == selected


def test_dev_virtualenv_requirement_excludes_affected_environment_builders() -> None:
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    requirements = [Requirement(value) for value in project["optional-dependencies"]["dev"]]
    virtualenv = next(requirement for requirement in requirements if requirement.name.lower() == "virtualenv")

    assert "21.14.7" not in virtualenv.specifier
    assert "21.14.8" not in virtualenv.specifier
    assert "21.14.9" in virtualenv.specifier
    assert "21.14.10" in virtualenv.specifier
