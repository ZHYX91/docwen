"""Layout plugin import and manifest contract tests."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


def test_plugin_layout_importable() -> None:
    import docwen_plugin_layout

    assert docwen_plugin_layout.__version__ == "0.1.0"


def test_layout_plugin_can_be_instantiated() -> None:
    from docwen_plugin_layout import LayoutPlugin

    plugin = LayoutPlugin()
    assert plugin.plugin_id == "docwen_plugin_layout"


def test_layout_plugin_manifest_declares_routes() -> None:
    from docwen_plugin_layout import LayoutPlugin

    manifest = LayoutPlugin().manifest
    assert manifest.plugin_id == "docwen_plugin_layout"
    routes = {(r.source_format, r.target_format, r.action_name) for r in manifest.routes}

    # ── Every fixed-layout source → every target ──────────────────────
    sources = ("pdf", "ofd", "xps")
    image_targets = ("png", "jpg", "tif")
    document_targets = ("docx", "doc", "odt", "rtf")

    for src in sources:
        # → Markdown
        assert (src, "md", "") in routes, f"missing: {src}→md"
        # → Images
        for tgt in image_targets:
            assert (src, tgt, "") in routes, f"missing: {src}→{tgt}"
        # → Documents (Office bridge)
        for tgt in document_targets:
            assert (src, tgt, "") in routes, f"missing: {src}→{tgt}"
        # → PDF (normalize / passthrough)
        assert (src, "pdf", "") in routes, f"missing: {src}→pdf"

    # ── Action routes ─────────────────────────────────────────────────
    assert ("pdf", "pdf", "merge_pdfs") in routes
    assert ("pdf", "pdf", "split_pdf") in routes

    # ── EPUB → MD is no longer in layout (moved to markup) ────────────
    assert ("epub", "md", "") not in routes

    # ── Unsupported or abstract sources are not registered ────────────
    for source in ("caj", "oxps", "layout"):
        assert (source, "md", "") not in routes
        assert (source, "png", "") not in routes

    # ── Total: 3 sources × 9 targets + 2 actions = 29 ─────────────────
    assert len(manifest.routes) == 29, f"Expected 29 routes, got {len(manifest.routes)}"

    assert manifest.requires == []


def test_split_pdf_manifest_declares_only_plugin_consumed_options() -> None:
    """GUI page-count context is local validation state, not a plugin request option."""
    from docwen_plugin_layout.manifest import PDF_SPLIT_OPTIONS_SCHEMA

    properties = PDF_SPLIT_OPTIONS_SCHEMA["properties"]
    assert set(properties) == {"split_mode", "pages"}
    assert properties["split_mode"]["enum"] == ["custom", "every_page", "odd_even"]
    assert "total_pages" not in properties


def test_layout_manifest_keeps_route_specific_options_scoped() -> None:
    """Layout route schemas should stay scoped to each target/action family."""
    from docwen_plugin_layout import LayoutPlugin

    route_map = {
        (route.source_format, route.target_format, route.action_name): route for route in LayoutPlugin().manifest.routes
    }

    sources = {"pdf", "ofd", "xps"}
    markdown_options = {
        "to_md_keep_images",
        "to_md_enable_ocr",
        "ocr_language",
        "locale",
        "yaml_key_labels",
        "image_mode",
        "image_link_style",
        "render_dpi",
    }
    image_options = {"render_dpi"}
    split_options = {"split_mode", "pages"}

    for source in sources:
        md_properties = route_map[(source, "md", "")].options_schema["properties"]
        assert set(md_properties) == markdown_options
        assert md_properties["image_mode"]["enum"] == ["file", "base64", "embed", "omit"]
        assert md_properties["image_link_style"]["enum"] == [
            "wiki_embed",
            "wiki_link",
            "markdown_embed",
            "markdown_link",
        ]
        assert md_properties["render_dpi"]["default"] == 200
        for key in ("locale", "yaml_key_labels"):
            assert md_properties[key]["x-docwen-status"] == "implemented"

        for target in {"png", "jpg", "tif"}:
            image_properties = route_map[(source, target, "")].options_schema["properties"]
            assert set(image_properties) == image_options
            assert image_properties["render_dpi"]["default"] == 150
            assert not ((markdown_options - image_options) & set(image_properties))

        for target in {"docx", "doc", "odt", "rtf"}:
            assert route_map[(source, target, "")].options_schema["properties"] == {}

        assert route_map[(source, "pdf", "")].options_schema["properties"] == {}

    assert route_map[("pdf", "pdf", "merge_pdfs")].options_schema["properties"] == {}
    split_properties = route_map[("pdf", "pdf", "split_pdf")].options_schema["properties"]
    assert set(split_properties) == split_options
    assert split_properties["split_mode"]["enum"] == ["custom", "every_page", "odd_even"]
    assert not (markdown_options & set(split_properties))


def test_layout_docstrings_track_only_supported_sources() -> None:
    """Layout docstrings must describe reachable fixed-layout sources only."""
    import docwen_plugin_layout
    import docwen_plugin_layout.to_document as to_document
    import docwen_plugin_layout.to_markdown as to_markdown
    from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter
    from docwen_plugin_layout.to_pdf.converter import LayoutToPdfConverter

    package_doc = docwen_plugin_layout.__doc__ or ""
    markdown_package_doc = to_markdown.__doc__ or ""
    markdown_doc = LayoutToMarkdownConverter.__doc__ or ""
    document_doc = to_document.__doc__ or ""
    pdf_doc = LayoutToPdfConverter.__doc__ or ""

    assert "PDF/OFD/XPS" in package_doc
    assert "PDF/OFD/XPS/layout" not in package_doc
    assert "Office bridges" in document_doc
    assert "PDF->DOCX fallback" in document_doc

    for doc in (package_doc, markdown_package_doc, markdown_doc, document_doc, pdf_doc):
        assert "CAJ" not in doc
        assert "OXPS" not in doc
    assert "NOT_IMPLEMENTED" not in document_doc


def test_layout_plugin_can_handle_explicit_routes() -> None:
    """can_handle() must return True for all explicitly registered routes."""
    from docwen_plugin_layout import LayoutPlugin

    plugin = LayoutPlugin()

    sources = ("pdf", "ofd", "xps")
    image_targets = ("png", "jpg", "tif")
    document_targets = ("docx", "doc", "odt", "rtf")

    for src in sources:
        # → Markdown
        assert plugin.can_handle(src, "md") is True, f"can_handle({src}, md) failed"
        # → Images
        for tgt in image_targets:
            assert plugin.can_handle(src, tgt) is True, f"can_handle({src}, {tgt}) failed"
        # → Documents
        for tgt in document_targets:
            assert plugin.can_handle(src, tgt) is True, f"can_handle({src}, {tgt}) failed"
        # → PDF
        assert plugin.can_handle(src, "pdf") is True, f"can_handle({src}, pdf) failed"

    # ── Action routes ─────────────────────────────────────────────────
    assert plugin.can_handle("pdf", "pdf", "merge_pdfs") is True
    assert plugin.can_handle("pdf", "pdf", "split_pdf") is True

    # ── EPUB is not in layout anymore ─────────────────────────────────
    assert plugin.can_handle("epub", "md") is False

    # ── Unsupported or abstract sources are not registered ────────────
    for source in ("caj", "oxps", "layout"):
        assert plugin.can_handle(source, "md") is False
        assert plugin.can_handle(source, "pdf") is False

    # ── Negative cases ────────────────────────────────────────────────
    assert plugin.can_handle("document", "docx") is False
    assert plugin.can_handle("image", "png") is False


def test_layout_plugin_only_depends_on_core() -> None:
    import sys

    import docwen_plugin_layout

    assert docwen_plugin_layout  # consumed by sys.modules inspection below

    forbidden = {
        "docwen_runtime",
        "docwen_application",
        "docwen_gui",
        "docwen_cli",
        "docwen_bundle",
        "docwen_plugin_document",
        "docwen_plugin_presentation",
        "docwen_plugin_markup",
        "docwen_plugin_spreadsheet",
        "docwen_plugin_image",
        "docwen_plugin_print",
        "docwen_plugin_markdown",
        "docwen_plugin_proofread",
        "docwen.",
    }
    plugin_modules = {name for name in sys.modules if name.startswith("docwen_plugin_layout")}
    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            for forbidden_pkg in forbidden:
                assert not val_mod.startswith(forbidden_pkg), f"{mod_name}.{attr} depends on {val_mod}"
