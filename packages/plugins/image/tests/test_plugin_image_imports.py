"""Image plugin import and manifest contract tests."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


def test_plugin_image_importable() -> None:
    import docwen_plugin_image

    assert docwen_plugin_image.__version__ == "0.1.0"


def test_image_plugin_can_be_instantiated() -> None:
    from docwen_plugin_image import ImagePlugin

    plugin = ImagePlugin()
    assert plugin.plugin_id == "docwen_plugin_image"


def test_image_plugin_manifest_declares_routes() -> None:
    from docwen_plugin_image import ImagePlugin

    manifest = ImagePlugin().manifest
    assert manifest.plugin_id == "docwen_plugin_image"
    routes = {(r.source_format, r.target_format, r.action_name) for r in manifest.routes}
    assert ("image", "md", "") in routes
    assert ("image", "pdf", "") in routes
    assert ("image", "image", "") in routes
    assert ("image", "tif", "merge_images_to_tiff") in routes
    assert manifest.requires == []


def test_image_plugin_manifest_uses_public_ocr_placement_values() -> None:
    """Reserved OCR placement schema must use the same public values as CLI/config."""
    from docwen_plugin_image.manifest import IMAGE_TO_MD_OPTIONS_SCHEMA

    schema = IMAGE_TO_MD_OPTIONS_SCHEMA["properties"]["ocr_placement"]
    assert schema["enum"] == ["image_md", "main_md"]
    assert "text_block" not in schema["enum"]


def test_image_to_md_manifest_declares_image_link_style_option() -> None:
    from docwen_plugin_image.manifest import IMAGE_TO_MD_OPTIONS_SCHEMA

    schema = IMAGE_TO_MD_OPTIONS_SCHEMA["properties"]["image_link_style"]
    assert schema["enum"] == ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"]
    assert schema["default"] == "wiki_embed"


def test_image_to_pdf_manifest_declares_only_consumed_quality_mode() -> None:
    from docwen_plugin_image.manifest import IMAGE_TO_PDF_OPTIONS_SCHEMA

    properties = IMAGE_TO_PDF_OPTIONS_SCHEMA["properties"]
    assert set(properties) == {"quality_mode"}
    assert properties["quality_mode"]["enum"] == ["original", "a4", "a3"]


def test_generic_image_format_manifest_declares_target_format_without_leaking_to_explicit_routes() -> None:
    from docwen_plugin_image.manifest import IMAGE_FORMAT_OPTIONS_SCHEMA, build_manifest

    properties = IMAGE_FORMAT_OPTIONS_SCHEMA["properties"]
    assert properties["target_format"]["enum"] == ["jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp"]
    assert properties["target_format"]["default"] == "png"

    route_map = {(r.source_format, r.target_format, r.action_name): r for r in build_manifest().routes}
    generic_schema = route_map[("image", "image", "")].options_schema["properties"]
    explicit_schema = route_map[("image", "webp", "")].options_schema["properties"]
    assert "target_format" in generic_schema
    assert "target_format" not in explicit_schema
    assert explicit_schema["compress_mode"]["enum"] == ["lossless", "limit_size"]
    assert explicit_schema["size_unit"]["enum"] == ["KB", "MB"]


def test_merge_images_to_tiff_manifest_declares_only_consumed_options() -> None:
    from docwen_plugin_image.manifest import IMAGE_MERGE_OPTIONS_SCHEMA

    properties = IMAGE_MERGE_OPTIONS_SCHEMA["properties"]
    assert set(properties) == {"mode", "keep_alpha"}
    assert properties["mode"]["enum"] == ["smart", "rgb", "RGB"]
    assert properties["keep_alpha"]["type"] == "boolean"


def test_image_manifest_keeps_route_specific_options_scoped() -> None:
    """Image route schemas should declare only the options their converters consume."""
    from docwen_plugin_image.manifest import build_manifest

    route_map = {
        (route.source_format, route.target_format, route.action_name): route for route in build_manifest().routes
    }

    markdown_options = {
        "to_md_keep_images",
        "to_md_enable_ocr",
        "ocr_language",
        "locale",
        "yaml_key_labels",
        "image_mode",
        "ocr_placement",
        "image_link_style",
    }
    image_format_options = {
        "compress_mode",
        "size_limit",
        "size_unit",
    }
    image_pdf_options = {"quality_mode"}
    merge_tiff_options = {"mode", "keep_alpha"}

    md_properties = route_map[("image", "md", "")].options_schema["properties"]
    assert set(md_properties) == markdown_options

    pdf_properties = route_map[("image", "pdf", "")].options_schema["properties"]
    assert set(pdf_properties) == image_pdf_options
    assert not (markdown_options & set(pdf_properties))

    generic_image_properties = route_map[("image", "image", "")].options_schema["properties"]
    assert set(generic_image_properties) == image_format_options | {"target_format"}

    for target in {"jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp"}:
        explicit_properties = route_map[("image", target, "")].options_schema["properties"]
        assert set(explicit_properties) == image_format_options
        assert "target_format" not in explicit_properties
        assert not (markdown_options & set(explicit_properties))

    merge_properties = route_map[("image", "tif", "merge_images_to_tiff")].options_schema["properties"]
    assert set(merge_properties) == merge_tiff_options
    assert not (markdown_options & set(merge_properties))
    assert "target_format" not in merge_properties
    assert "pdf_quality" not in merge_properties
    assert "compress_mode" not in merge_properties


def test_image_plugin_can_handle_manifest_routes() -> None:
    from docwen_plugin_image import ImagePlugin

    plugin = ImagePlugin()
    assert plugin.can_handle("image", "md") is True
    assert plugin.can_handle("image", "pdf") is True
    assert plugin.can_handle("image", "png") is True
    assert plugin.can_handle("image", "tif", "merge_images_to_tiff") is True
    assert plugin.can_handle("image", "pdf", "merge_images_to_tiff") is False
    assert plugin.can_handle("document", "md") is False


def test_image_plugin_can_handle_subformat_source() -> None:
    """M-3: can_handle should work when source_format is a concrete format
    rather than the category ('image'). The runtime route resolver is
    expected to normalise concrete formats to categories before calling
    can_handle, but this test serves as a contract baseline."""
    from docwen_plugin_image import ImagePlugin

    plugin = ImagePlugin()
    # Concrete sub-format as source → returned by manifest's explicit routes
    assert plugin.can_handle("image", "jpg") is True
    assert plugin.can_handle("image", "png") is True
    # These are NOT expected to match (the manifest's RouteSpec uses source_format="image", not "jpg")
    # Documenting the current contract: route resolver must normalise jpg→image before calling can_handle
    assert plugin.can_handle("jpg", "png") is False, "route resolver must normalise jpg→image before can_handle"
    assert plugin.can_handle("png", "jpg") is False, "route resolver must normalise png→image before can_handle"
    assert plugin.can_handle("png", "tif") is False


def test_image_plugin_only_depends_on_core() -> None:
    import sys

    import docwen_plugin_image

    assert docwen_plugin_image  # consumed by sys.modules inspection below

    forbidden = {
        "docwen_runtime",
        "docwen_application",
        "docwen_gui",
        "docwen_cli",
        "docwen_bundle",
        "docwen_plugin_document",
        "docwen_plugin_spreadsheet",
        "docwen.",
    }
    plugin_modules = {name for name in sys.modules if name.startswith("docwen_plugin_image")}
    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            for forbidden_pkg in forbidden:
                assert not val_mod.startswith(forbidden_pkg), f"{mod_name}.{attr} depends on {val_mod}"
