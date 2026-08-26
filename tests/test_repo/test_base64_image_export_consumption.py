from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]

SHARED_MARKDOWN_CONSUMERS = (
    ROOT / "packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/converter.py",
    ROOT / "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/to_markdown/converter.py",
    ROOT / "packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py",
    ROOT / "packages/plugins/layout/src/docwen_plugin_layout/to_markdown/converter.py",
)
DIRECT_DATA_URI_CONSUMERS = (
    ROOT / "packages/plugins/image/src/docwen_plugin_image/to_markdown/converter.py",
    ROOT / "packages/plugins/markup/src/docwen_plugin_markup/markdown_resources.py",
)


def test_base64_export_config_has_one_declared_owner_and_gui_persistence_path() -> None:
    conversion = tomllib.loads((ROOT / "configs/conversion.toml").read_text(encoding="utf-8"))
    export = tomllib.loads((ROOT / "configs/export.toml").read_text(encoding="utf-8"))
    settings_vm = (ROOT / "packages/apps/gui/src/docwen_gui/view_models/settings_vm.py").read_text(encoding="utf-8")

    assert conversion["export"] == {
        "base64_compress_enabled": True,
        "base64_compress_threshold_kb": 100,
    }
    assert "base64_compress_enabled" not in export
    assert "base64_compress_threshold_kb" not in export
    assert 'conv_export.get("base64_compress_enabled", True)' in settings_vm
    assert 'conv_export.get("base64_compress_threshold_kb", 100)' in settings_vm
    assert 'put("conversion.export.base64_compress_enabled"' in settings_vm
    assert 'put("conversion.export.base64_compress_threshold_kb"' in settings_vm


def test_core_owns_the_configured_base64_image_encoder_and_declares_pillow() -> None:
    core_project = tomllib.loads((ROOT / "packages/core/pyproject.toml").read_text(encoding="utf-8"))
    dependencies = core_project["project"]["dependencies"]
    semantics_source = (ROOT / "packages/core/src/docwen_core/export_semantics/__init__.py").read_text(encoding="utf-8")
    encoder_source = (ROOT / "packages/core/src/docwen_core/text/image_markdown.py").read_text(encoding="utf-8")

    assert any(dependency.lower().startswith("pillow") for dependency in dependencies)
    assert 'conversion_export = conv.get("export", {})' in semantics_source
    assert 'conversion_export.get("base64_compress_enabled", True)' in semantics_source
    assert 'conversion_export.get("base64_compress_threshold_kb", 100)' in semantics_source
    assert "export_semantics: MarkdownExportSemantics," in encoder_source
    assert "never read process-global configuration" in encoder_source
    assert "export_semantics: MarkdownExportSemantics | None = None" in encoder_source
    assert "if export_semantics is None:" in encoder_source
    assert "compress_enabled = export_semantics.export_base64_compress_enabled" in encoder_source
    assert "from docwen_core.formats import CATEGORY_IMAGE, get_category, get_media_type" in encoder_source
    assert "len(source_bytes) > threshold_bytes" in encoder_source
    assert 'resolved_media_type = "image/jpeg"' in encoder_source


def test_all_six_markdown_plugin_families_use_the_shared_base64_encoder() -> None:
    for consumer in SHARED_MARKDOWN_CONSUMERS:
        source = consumer.read_text(encoding="utf-8")
        assert "generate_image_markdown" in source, consumer
        assert "base64.b64encode" not in source, consumer

    for consumer in DIRECT_DATA_URI_CONSUMERS:
        source = consumer.read_text(encoding="utf-8")
        assert "build_base64_image_data_uri" in source, consumer
        assert "base64.b64encode" not in source, consumer
