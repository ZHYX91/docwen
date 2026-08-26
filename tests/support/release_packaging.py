"""Shared filesystem fixtures for release verifier tests."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

_PACKAGED_LAYOUT_FIXTURE_BYTES = b"packaged-resource"


def use_compact_pymupdf_layout_manifest(monkeypatch: pytest.MonkeyPatch) -> None:
    """Keep verifier unit fixtures small while exercising the real hash path."""

    from docwen_runtime import pymupdf_layout_resources

    manifest = (
        pymupdf_layout_resources.PymupdfLayoutResourceSpec(
            relative_path="onnx/test-layout.onnx",
            size=len(_PACKAGED_LAYOUT_FIXTURE_BYTES),
            sha256=hashlib.sha256(_PACKAGED_LAYOUT_FIXTURE_BYTES).hexdigest(),
        ),
    )
    monkeypatch.setattr(
        pymupdf_layout_resources,
        "_current_pymupdf_layout_resource_manifest",
        lambda: manifest,
    )


def write_packaged_common_resources(binary_dir: Path) -> None:
    """Populate the common resource tree required by both packaged verifiers."""

    from scripts.release import packaged_resources, verify_packaged_cli

    for rel_path in verify_packaged_cli._REQUIRED_CONFIG_FILES:
        config_path = binary_dir / "configs" / rel_path
        config_path.parent.mkdir(parents=True, exist_ok=True)
        config_path.write_text("[placeholder]\n", encoding="utf-8")
    for rel_path in verify_packaged_cli._REQUIRED_TEMPLATE_FILES:
        template_path = binary_dir / "templates" / rel_path
        template_path.parent.mkdir(parents=True, exist_ok=True)
        template_path.write_bytes(b"template")
    for rel_path in verify_packaged_cli._REQUIRED_MODEL_FILES:
        model_path = binary_dir / "models" / rel_path
        model_path.parent.mkdir(parents=True, exist_ok=True)
        model_path.write_bytes(b"onnx")
    locales_dir = binary_dir / "_internal" / "docwen" / "i18n" / "locales"
    locales_dir.mkdir(parents=True)
    for rel_path in verify_packaged_cli._REQUIRED_LOCALE_FILES:
        (locales_dir / rel_path).write_text("[meta]\n", encoding="utf-8")
    pymupdf_layout_dir = binary_dir / packaged_resources.PYMUPDF_LAYOUT_PACKAGED_RESOURCE_ROOT
    for rel_path in packaged_resources.pymupdf_layout_resource_paths():
        resource_path = pymupdf_layout_dir / rel_path
        resource_path.parent.mkdir(parents=True, exist_ok=True)
        resource_path.write_bytes(_PACKAGED_LAYOUT_FIXTURE_BYTES)


def write_packaged_gui_assets(binary_dir: Path) -> None:
    """Populate the GUI-only asset tree required by its packaged verifier."""

    from scripts.release import verify_packaged_gui

    for rel_path in verify_packaged_gui._REQUIRED_ASSET_FILES:
        asset_path = binary_dir / "assets" / rel_path
        asset_path.parent.mkdir(parents=True, exist_ok=True)
        asset_path.write_bytes(b"asset")


__all__ = [
    "use_compact_pymupdf_layout_manifest",
    "write_packaged_common_resources",
    "write_packaged_gui_assets",
]
