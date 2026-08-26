"""Markup plugin import and manifest contract tests."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


def test_plugin_markup_importable() -> None:
    import docwen_plugin_markup

    assert docwen_plugin_markup.__version__ == "0.1.0"


def test_markup_plugin_can_be_instantiated() -> None:
    from docwen_plugin_markup import MarkupPlugin

    plugin = MarkupPlugin()
    assert plugin.plugin_id == "docwen_plugin_markup"


def test_markup_plugin_manifest_declares_routes() -> None:
    from docwen_plugin_markup import MarkupPlugin

    manifest = MarkupPlugin().manifest
    assert manifest.plugin_id == "docwen_plugin_markup"
    routes = {(r.source_format, r.target_format, r.action_name) for r in manifest.routes}

    # ── Web archive → MD ──────────────────────────────────────────────
    assert ("html", "md", "") in routes
    assert ("htm", "md", "") in routes
    assert ("mhtml", "md", "") in routes
    assert ("mht", "md", "") in routes

    # ── Note export → MD ──────────────────────────────────────────────
    assert ("enex", "md", "") in routes

    # ── Publication → MD ──────────────────────────────────────────────
    assert ("epub", "md", "") in routes

    # ── Total: 6 routes ───────────────────────────────────────────────
    assert len(manifest.routes) == 6, f"Expected 6 routes, got {len(manifest.routes)}"

    assert manifest.requires == []


def test_markup_to_md_manifest_declares_image_link_style_option() -> None:
    from docwen_plugin_markup import MarkupPlugin

    manifest = MarkupPlugin().manifest
    markdown_routes = [
        route
        for route in manifest.routes
        if route.target_format == "md" and route.source_format in {"html", "htm", "mhtml", "mht", "enex", "epub"}
    ]
    assert len(markdown_routes) == 6

    for route in markdown_routes:
        properties = route.options_schema["properties"]
        schema = properties["image_link_style"]
        assert schema["enum"] == ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"]
        assert schema["default"] == "wiki_embed"


def test_markup_to_md_manifest_declares_consumed_markdown_export_options() -> None:
    """Markup-family -> MD routes should expose all consumed Markdown export options."""
    from docwen_plugin_markup import MarkupPlugin

    route_map = {
        (route.source_format, route.target_format, route.action_name): route for route in MarkupPlugin().manifest.routes
    }

    consumed_options = {
        "to_md_keep_images",
        "to_md_enable_ocr",
        "ocr_language",
        "locale",
        "yaml_key_labels",
        "image_mode",
        "ocr_placement",
        "image_link_style",
    }
    markup_sources = {"html", "htm", "mhtml", "mht", "enex", "epub"}

    assert set(route_map) == {(source, "md", "") for source in markup_sources}
    for source in markup_sources:
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


def test_markup_plugin_can_handle_routes() -> None:
    from docwen_plugin_markup import MarkupPlugin

    plugin = MarkupPlugin()

    # Web archive
    for fmt in ("html", "htm", "mhtml", "mht"):
        assert plugin.can_handle(fmt, "md") is True, f"can_handle({fmt}, md) failed"

    # Note export
    assert plugin.can_handle("enex", "md") is True

    # Publication
    assert plugin.can_handle("epub", "md") is True

    # The old target spelling is not an alias.
    assert plugin.can_handle("html", "markdown") is False

    # ── Negative cases ────────────────────────────────────────────────
    assert plugin.can_handle("docx", "md") is False
    assert plugin.can_handle("pdf", "md") is False
    assert plugin.can_handle("pptx", "md") is False
    assert plugin.can_handle("html", "pdf") is False


def test_markup_plugin_only_depends_on_core() -> None:
    import sys

    import docwen_plugin_markup  # type: ignore[unused-import]  # noqa: F401

    forbidden = {
        "docwen_runtime",
        "docwen_application",
        "docwen_gui",
        "docwen_cli",
        "docwen_bundle",
        "docwen_plugin_document",
        "docwen_plugin_presentation",
        "docwen_plugin_layout",
        "docwen_plugin_spreadsheet",
        "docwen_plugin_image",
        "docwen_plugin_markdown",
        "docwen_plugin_print",
        "docwen_plugin_proofread",
        "docwen.",
    }
    plugin_modules = {name for name in sys.modules if name.startswith("docwen_plugin_markup")}
    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            for forbidden_pkg in forbidden:
                assert not val_mod.startswith(forbidden_pkg), f"{mod_name}.{attr} depends on {val_mod}"
