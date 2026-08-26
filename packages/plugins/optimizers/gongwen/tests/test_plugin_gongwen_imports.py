"""Tests for Gongwen plugin import and public option contracts."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


_OCR_LANGUAGE_ENUM = ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"]


def test_gongwen_plugin_importable() -> None:
    import docwen_plugin_optimizer_gongwen

    assert docwen_plugin_optimizer_gongwen.__version__ == "0.1.0"


def test_gongwen_manifest_declares_route_and_ocr_options() -> None:
    from docwen_plugin_optimizer_gongwen import GongwenOptimizerPlugin

    plugin = GongwenOptimizerPlugin()
    manifest = plugin.manifest

    assert manifest.plugin_id == "docwen_plugin_optimizer_gongwen"
    routes = {(route.source_format, route.target_format, route.action_name) for route in manifest.routes}
    assert ("docx", "md", "gongwen") in routes
    assert [resource.to_dict() for resource in manifest.optimization_resources] == [
        {
            "id": "gongwen",
            "name": "Chinese official-document optimization",
            "action_name": "gongwen",
        }
    ]
    assert "supported_actions" not in manifest.extra

    route = manifest.routes[0]
    props = route.options_schema["properties"]
    assert props["to_md_enable_ocr"]["type"] == "boolean"
    assert props["to_md_enable_ocr"]["default"] is False
    assert props["ocr_language"]["enum"] == _OCR_LANGUAGE_ENUM
    assert props["ocr_language"]["default"] == "auto"


def test_gongwen_manifest_declares_only_consumed_action_options() -> None:
    """Gongwen schema should expose public request options, not internal paths."""
    from docwen_plugin_optimizer_gongwen import GongwenOptimizerPlugin

    route = GongwenOptimizerPlugin().manifest.routes[0]
    props = route.options_schema["properties"]

    assert (route.source_format, route.target_format, route.action_name) == ("docx", "md", "gongwen")
    assert set(props) == {
        "remove_numbering",
        "add_numbering",
        "numbering_scheme",
        "to_md_keep_images",
        "image_mode",
        "image_link_style",
        "to_md_enable_ocr",
        "ocr_language",
        "locale",
        "table_merge_strategy",
    }
    assert props["numbering_scheme"]["enum"] == [
        "gongwen_standard",
        "hierarchical_standard",
        "hierarchical_h2_start",
        "legal_standard",
    ]
    assert props["ocr_language"]["enum"] == _OCR_LANGUAGE_ENUM
    assert props["locale"]["x-docwen-status"] == "implemented"
    assert props["table_merge_strategy"]["enum"] == ["fill", "empty", "marker"]
    assert props["to_md_keep_images"]["default"] is True
    assert props["image_mode"]["enum"] == ["file", "base64", "embed", "omit"]
    assert props["image_link_style"]["enum"] == [
        "wiki_embed",
        "wiki_link",
        "markdown_embed",
        "markdown_link",
    ]
    assert "yaml_key_labels" not in props
    assert "output_dir" not in props
