"""Repo checks that packaged resource manifests match their source trees."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _config_registry_paths() -> set[str]:
    registry_path = Path("packages/runtime/src/docwen_runtime/config/registry.py")
    tree = ast.parse(registry_path.read_text(encoding="utf-8"))
    paths: set[str] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Name) or node.func.id != "ConfigFileSpec":
            continue
        if not node.args or not isinstance(node.args[0], ast.Constant) or not isinstance(node.args[0].value, str):
            continue
        paths.add(node.args[0].value)
    return paths


def test_packaged_verifiers_share_common_resource_manifest_source() -> None:
    cli_source = Path("scripts/release/verify_packaged_cli.py").read_text(encoding="utf-8")
    gui_source = Path("scripts/release/verify_packaged_gui.py").read_text(encoding="utf-8")
    shared_source = Path("scripts/release/packaged_resources.py").read_text(encoding="utf-8")

    assert "verify_common_resource_layout" in cli_source
    assert "verify_common_resource_layout" in gui_source
    assert "REQUIRED_CONFIG_FILES = (" in shared_source
    assert "REQUIRED_TEMPLATE_FILES = (" in shared_source
    assert "REQUIRED_MODEL_FILES = (" in shared_source
    assert "REQUIRED_LOCALE_FILES = (" in shared_source
    assert "REQUIRED_ASSET_FILES = (" in shared_source
    assert "verify_no_bundled_qt_network_stack" in shared_source
    assert "_REQUIRED_CONFIG_FILES = (" not in cli_source
    assert "_REQUIRED_CONFIG_FILES = (" not in gui_source
    assert "_REQUIRED_TEMPLATE_FILES = (" not in cli_source
    assert "_REQUIRED_TEMPLATE_FILES = (" not in gui_source
    assert "_REQUIRED_MODEL_FILES = (" not in cli_source
    assert "_REQUIRED_MODEL_FILES = (" not in gui_source
    assert "_REQUIRED_LOCALE_FILES = (" not in cli_source
    assert "_REQUIRED_LOCALE_FILES = (" not in gui_source
    assert "_REQUIRED_ASSET_FILES = (" not in gui_source


@pytest.mark.parametrize(
    "relative_path",
    [
        "_internal/PySide6/Qt6Network.dll",
        "_internal/PySide6/QtNetwork.pyd",
        "_internal/PySide6/plugins/networkinformation/qnetworklistmanager.dll",
        "_internal/PySide6/plugins/tls/qschannelbackend.dll",
        "_internal/PySide6/Qt/lib/QtNetwork.framework/Versions/A/QtNetwork",
        "_internal/PySide6/Qt/lib/libQt6Network.so.6",
    ],
)
def test_packaged_resource_gate_rejects_qt_network_stack(tmp_path: Path, relative_path: str) -> None:
    from scripts.release.packaged_resources import verify_no_bundled_qt_network_stack

    target = tmp_path / relative_path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(b"forbidden")

    with pytest.raises(RuntimeError, match="forbidden_qt_network_stack"):
        verify_no_bundled_qt_network_stack(tmp_path, error_prefix="test")


def test_packaged_model_manifests_match_current_runtime_models() -> None:
    from scripts.release import verify_packaged_cli, verify_packaged_gui

    models_dir = Path("models")
    runtime_models = {path.relative_to(models_dir).as_posix() for path in models_dir.rglob("*.onnx") if path.is_file()}

    assert set(verify_packaged_cli._REQUIRED_MODEL_FILES) == runtime_models
    assert set(verify_packaged_gui._REQUIRED_MODEL_FILES) == runtime_models


def test_packaged_ocr_model_manifest_covers_configured_language_models() -> None:
    from scripts.release import packaged_resources

    from docwen_core.text import OCR_LANGUAGE_MODELS

    configured_model_files = {
        f"rapidocr/{filename}" for model_files in OCR_LANGUAGE_MODELS.values() for filename in model_files.values()
    }
    required_model_files = set(packaged_resources.REQUIRED_MODEL_FILES)
    reserved_model_files = required_model_files - configured_model_files

    assert configured_model_files <= required_model_files
    assert reserved_model_files == {"rapidocr/arabic_PP-OCRv4_rec_infer.onnx"}
    assert "arabic" not in OCR_LANGUAGE_MODELS
    assert Path("models/rapidocr/arabic_PP-OCRv4_rec_infer.onnx").is_file()


def test_packaged_template_manifests_match_current_runtime_templates() -> None:
    from scripts.release import verify_packaged_cli, verify_packaged_gui

    templates_dir = Path("templates")
    runtime_templates = {
        path.name for path in templates_dir.iterdir() if path.is_file() and path.suffix.lower() in {".docx", ".xlsx"}
    }

    assert set(verify_packaged_cli._REQUIRED_TEMPLATE_FILES) == runtime_templates
    assert set(verify_packaged_gui._REQUIRED_TEMPLATE_FILES) == runtime_templates


def test_packaged_config_manifests_match_current_runtime_config_registry() -> None:
    from scripts.release import verify_packaged_cli, verify_packaged_gui

    registry_configs = _config_registry_paths()
    source_configs = {
        path.relative_to("configs").as_posix() for path in Path("configs").rglob("*.toml") if path.is_file()
    }

    assert set(verify_packaged_cli._REQUIRED_CONFIG_FILES) == registry_configs
    assert set(verify_packaged_gui._REQUIRED_CONFIG_FILES) == registry_configs
    assert registry_configs == source_configs


def test_packaged_locale_manifests_match_current_runtime_locales() -> None:
    from scripts.release import verify_packaged_cli, verify_packaged_gui

    locales_dir = Path("i18n/locales")
    runtime_locales = {path.name for path in locales_dir.glob("*.toml") if path.is_file()}

    assert set(verify_packaged_cli._REQUIRED_LOCALE_FILES) == runtime_locales
    assert set(verify_packaged_gui._REQUIRED_LOCALE_FILES) == runtime_locales
