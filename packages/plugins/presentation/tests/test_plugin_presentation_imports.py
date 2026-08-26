"""Presentation plugin import and manifest contract tests."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


def test_plugin_presentation_importable() -> None:
    import docwen_plugin_presentation

    assert docwen_plugin_presentation.__version__ == "0.1.0"


def test_presentation_plugin_can_be_instantiated() -> None:
    from docwen_plugin_presentation import PresentationPlugin

    plugin = PresentationPlugin()
    assert plugin.plugin_id == "docwen_plugin_presentation"


def test_presentation_plugin_manifest_declares_routes() -> None:
    from docwen_plugin_presentation import PresentationPlugin

    manifest = PresentationPlugin().manifest
    assert manifest.plugin_id == "docwen_plugin_presentation"
    routes = {(r.source_format, r.target_format, r.action_name) for r in manifest.routes}

    # ── PPTX → MD ─────────────────────────────────────────────────────
    assert ("pptx", "md", "") in routes

    # ── PPT → MD ──────────────────────────────────────────────────────
    assert ("ppt", "md", "") in routes

    # ── Total: 2 routes ───────────────────────────────────────────────
    assert len(manifest.routes) == 2, f"Expected 2 routes, got {len(manifest.routes)}"

    assert manifest.requires == []


def test_presentation_manifest_declares_image_mode_option() -> None:
    from docwen_plugin_presentation import PresentationPlugin

    for route in PresentationPlugin().manifest.routes:
        properties = route.options_schema["properties"]
        assert properties["image_mode"]["enum"] == ["file", "base64", "embed", "omit"]
        assert properties["image_mode"]["default"] == "file"


def test_presentation_to_md_manifest_declares_consumed_markdown_export_options() -> None:
    """PPT/PPTX -> MD routes should expose all consumed Markdown export options."""
    from docwen_plugin_presentation import PresentationPlugin

    route_map = {
        (route.source_format, route.target_format, route.action_name): route
        for route in PresentationPlugin().manifest.routes
    }

    consumed_options = {
        "export_notes",
        "to_md_keep_images",
        "to_md_enable_ocr",
        "ocr_language",
        "locale",
        "yaml_key_labels",
        "image_mode",
        "ocr_placement",
        "image_link_style",
    }

    for source in {"pptx", "ppt"}:
        properties = route_map[(source, "md", "")].options_schema["properties"]
        assert set(properties) == consumed_options
        assert properties["image_mode"]["enum"] == ["file", "base64", "embed", "omit"]
        assert properties["ocr_placement"]["enum"] == ["image_md", "main_md"]
        assert properties["image_link_style"]["enum"] == [
            "wiki_embed",
            "wiki_link",
            "markdown_embed",
            "markdown_link",
        ]
        for key in ("locale", "yaml_key_labels"):
            assert properties[key]["x-docwen-status"] == "implemented"


def test_presentation_plugin_can_handle_routes() -> None:
    from docwen_plugin_presentation import PresentationPlugin

    plugin = PresentationPlugin()

    assert plugin.can_handle("pptx", "md") is True
    assert plugin.can_handle("ppt", "md") is True

    # The old target spelling is not an alias.
    assert plugin.can_handle("pptx", "markdown") is False

    # ── Negative cases ────────────────────────────────────────────────
    assert plugin.can_handle("docx", "md") is False
    assert plugin.can_handle("pdf", "md") is False
    assert plugin.can_handle("html", "md") is False
    assert plugin.can_handle("pptx", "pdf") is False


def test_presentation_plugin_only_depends_on_core() -> None:
    import sys

    import docwen_plugin_presentation

    assert docwen_plugin_presentation  # consumed by sys.modules inspection below

    forbidden = {
        "docwen_runtime",
        "docwen_application",
        "docwen_gui",
        "docwen_cli",
        "docwen_bundle",
        "docwen_plugin_document",
        "docwen_plugin_markup",
        "docwen_plugin_layout",
        "docwen_plugin_spreadsheet",
        "docwen_plugin_image",
        "docwen_plugin_markdown",
        "docwen_plugin_print",
        "docwen_plugin_proofread",
        "docwen.",
    }
    plugin_modules = {name for name in sys.modules if name.startswith("docwen_plugin_presentation")}
    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            for forbidden_pkg in forbidden:
                assert not val_mod.startswith(forbidden_pkg), f"{mod_name}.{attr} depends on {val_mod}"
