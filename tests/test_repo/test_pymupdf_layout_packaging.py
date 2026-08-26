"""Contracts for packaging PyMuPDF Layout's split-distribution resources."""

from __future__ import annotations

import importlib.metadata
import shutil
import tomllib
from pathlib import Path, PurePosixPath

import pytest

pytestmark = pytest.mark.unit


def _write_packaged_layout_resources(binary_dir: Path, *, copy_real_bytes: bool) -> tuple[str, ...]:
    from scripts.release import packaged_resources

    from docwen_runtime.pymupdf_layout_resources import (
        PYMUPDF_LAYOUT_DISTRIBUTION,
        PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT,
        pymupdf_layout_resource_manifest,
    )

    resources = packaged_resources.pymupdf_layout_source_resource_files()
    resource_root = binary_dir / packaged_resources.PYMUPDF_LAYOUT_PACKAGED_RESOURCE_ROOT
    source_root = Path(
        str(
            importlib.metadata.distribution(PYMUPDF_LAYOUT_DISTRIBUTION).locate_file(
                PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT
            )
        )
    )
    specs = {spec.relative_path: spec for spec in pymupdf_layout_resource_manifest()}
    for index, name in enumerate(resources):
        destination = resource_root / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        if copy_real_bytes:
            shutil.copyfile(source_root / name, destination)
        elif index == 0:
            destination.write_bytes(bytes(specs[name].size))
        else:
            destination.write_bytes(b"x")
    return resources


def test_pymupdf_layout_manifest_tracks_current_distribution_resources() -> None:
    from scripts.release import packaged_resources

    from docwen_runtime.pymupdf_layout_resources import (
        PYMUPDF_LAYOUT_DISTRIBUTION,
        PYMUPDF_LAYOUT_DISTRIBUTION_VERSION,
    )

    resources = packaged_resources.pymupdf_layout_source_resource_files()
    lock = tomllib.loads(Path("uv.lock").read_text(encoding="utf-8"))
    locked_versions = [
        package["version"] for package in lock["package"] if package["name"] == PYMUPDF_LAYOUT_DISTRIBUTION
    ]

    assert resources == tuple(sorted(resources))
    assert resources
    assert PYMUPDF_LAYOUT_DISTRIBUTION_VERSION == "1.27.2.2"
    assert locked_versions == [PYMUPDF_LAYOUT_DISTRIBUTION_VERSION]
    assert len(resources) == 7
    assert all(not PurePosixPath(name).is_absolute() for name in resources)
    assert {PurePosixPath(name).suffix.lower() for name in resources} >= {".onnx", ".yaml"}
    assert not any("pymupdf_layout" in name for name in resources)


def test_packaged_layout_verifier_accepts_complete_distribution_manifest(tmp_path: Path) -> None:
    from scripts.release import packaged_resources

    _write_packaged_layout_resources(tmp_path, copy_real_bytes=True)

    packaged_resources.verify_pymupdf_layout_resource_layout(tmp_path, error_prefix="test")


def test_packaged_layout_verifier_fails_when_one_distribution_resource_is_missing(tmp_path: Path) -> None:
    from scripts.release import packaged_resources

    with pytest.raises(RuntimeError, match="test_pymupdf_layout_resources_invalid:required_resource_missing") as exc:
        packaged_resources.verify_pymupdf_layout_resource_layout(tmp_path, error_prefix="test")
    assert str(tmp_path) not in str(exc.value)


def test_packaged_layout_verifier_fails_when_one_distribution_resource_has_wrong_size(tmp_path: Path) -> None:
    from scripts.release import packaged_resources

    _write_packaged_layout_resources(tmp_path, copy_real_bytes=False)
    resource_root = tmp_path / packaged_resources.PYMUPDF_LAYOUT_PACKAGED_RESOURCE_ROOT
    first_resource = packaged_resources.pymupdf_layout_source_resource_files()[0]
    (resource_root / first_resource).write_bytes(b"x")

    with pytest.raises(RuntimeError, match="test_pymupdf_layout_resources_invalid:required_resource_size_mismatch"):
        packaged_resources.verify_pymupdf_layout_resource_layout(tmp_path, error_prefix="test")


def test_packaged_layout_verifier_fails_when_same_size_resource_hash_is_wrong(tmp_path: Path) -> None:
    from scripts.release import packaged_resources

    _write_packaged_layout_resources(tmp_path, copy_real_bytes=False)

    with pytest.raises(RuntimeError, match="test_pymupdf_layout_resources_invalid:required_resource_hash_mismatch"):
        packaged_resources.verify_pymupdf_layout_resource_layout(tmp_path, error_prefix="test")


def test_build_collects_the_real_pymupdf_layout_import_package() -> None:
    from scripts.build import build

    source = Path("scripts/build/build.py").read_text(encoding="utf-8")

    assert "pymupdf.layout" in build._PYINSTALLER_COMMON_COLLECT_ALL_TARGETS
    assert '"pymupdf_layout"' not in source
    assert "docwen.services.strategies" not in source
    assert "docwen.ipc" not in source
    assert build._pyinstaller_collection_args("collect-all", ("pymupdf.layout",)) == ["--collect-all=pymupdf.layout"]


def test_build_collection_preflight_accepts_current_package_targets() -> None:
    from scripts.build import build

    targets = (
        *build._PYINSTALLER_COMMON_COLLECT_ALL_TARGETS,
        *build._PYINSTALLER_GUI_COLLECT_ALL_TARGETS,
        *build._PYINSTALLER_COMMON_COLLECT_DATA_TARGETS,
        *build._PYINSTALLER_COMMON_COLLECT_SUBMODULE_TARGETS,
    )

    build._validate_pyinstaller_package_targets(targets)
    build._validate_pymupdf_layout_pyinstaller_data_collection()


def test_build_collection_preflight_fails_closed_for_invalid_target() -> None:
    from scripts.build import build

    with pytest.raises(RuntimeError, match="pyinstaller_collection_targets_invalid"):
        build._validate_pyinstaller_package_targets(("pymupdf_layout",))


def test_build_collection_preflight_fails_closed_when_pyinstaller_skips_layout_data(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PyInstaller.utils import hooks
    from scripts.build import build

    monkeypatch.setattr(hooks, "collect_data_files", lambda _package: [])

    with pytest.raises(RuntimeError, match="pyinstaller_pymupdf_layout_data_collection_incomplete"):
        build._validate_pymupdf_layout_pyinstaller_data_collection()
