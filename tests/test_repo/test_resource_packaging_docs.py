"""Current resource manifests and their documentation stay aligned."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]


def _config_registry_paths() -> set[str]:
    registry_path = ROOT / "packages" / "runtime" / "src" / "docwen_runtime" / "config" / "registry.py"
    tree = ast.parse(registry_path.read_text(encoding="utf-8"))
    paths: set[str] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Name) or node.func.id != "ConfigFileSpec":
            continue
        if node.args and isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str):
            paths.add(node.args[0].value)
    return paths


def _relative_files(root: Path, pattern: str = "*") -> set[str]:
    return {path.relative_to(root).as_posix() for path in root.rglob(pattern) if path.is_file()}


def test_packaged_resource_manifest_matches_the_current_source_tree() -> None:
    from scripts.release.packaged_resources import (
        REQUIRED_ASSET_FILES,
        REQUIRED_CONFIG_FILES,
        REQUIRED_LOCALE_FILES,
        REQUIRED_MODEL_FILES,
        REQUIRED_TEMPLATE_FILES,
    )

    assets = {path for path in _relative_files(ROOT / "assets") if not path.startswith("screenshots/")}
    templates = {
        path.name
        for path in (ROOT / "templates").iterdir()
        if path.is_file() and path.suffix.casefold() in {".docx", ".xlsx"}
    }

    assert set(REQUIRED_ASSET_FILES) == assets
    assert set(REQUIRED_CONFIG_FILES) == _config_registry_paths()
    assert set(REQUIRED_TEMPLATE_FILES) == templates
    assert set(REQUIRED_MODEL_FILES) == _relative_files(ROOT / "models", "*.onnx")
    assert set(REQUIRED_LOCALE_FILES) == {path.name for path in (ROOT / "i18n" / "locales").glob("*.toml")}


def test_build_script_collects_locales_from_the_canonical_source() -> None:
    source = (ROOT / "scripts" / "build" / "build.py").read_text(encoding="utf-8")

    assert 'i18n_locales_src = PROJECT_ROOT / "i18n" / "locales"' in source
    assert 'f"{i18n_locales_src}{os.pathsep}docwen/i18n/locales"' in source


def test_current_docs_describe_packaging_and_config_ownership() -> None:
    packaging = (ROOT / "docs" / "packaging.md").read_text(encoding="utf-8")
    configuration = (ROOT / "docs" / "configuration.md").read_text(encoding="utf-8")

    for resource in ("configs/", "templates/", "models/", "locale files", "application assets"):
        assert resource in packaging
    assert "Windows package resource/layout verification" in packaging
    assert "23 files" in configuration
    assert "configs/numbering/add.toml" in configuration
    assert "configs/numbering/cleanup.toml" in configuration
    assert "docwen_plugin_markdown.field_processors.gongwen.process_yaml" in configuration
    assert "docwen_plugin_markdown.template_filler" in configuration


def test_capability_inventory_tracks_current_resource_entrypoints() -> None:
    capabilities = (ROOT / "docs" / "capabilities.md").read_text(encoding="utf-8")

    required_tokens = (
        "| FEAT-RES-001 |",
        "| FEAT-RES-002 |",
        "| FEAT-RES-004 |",
        "| FEAT-RES-006 |",
        "| FEAT-RES-007 |",
        "scripts/release/packaged_resources.py",
        "docwen_bundle.cli_entry:main",
        "docwen_bundle.gui_entry:main",
    )
    for token in required_tokens:
        assert token in capabilities
