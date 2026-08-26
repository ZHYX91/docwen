"""Test that docwen_plugin_document can be imported and satisfies contracts."""

import pytest

pytestmark = pytest.mark.unit


def test_plugin_docx_importable() -> None:
    """docwen_plugin_document should be importable with version."""
    import docwen_plugin_document

    assert docwen_plugin_document.__version__ == "0.1.0"


def test_docx_plugin_can_be_instantiated() -> None:
    """DocumentPlugin should be importable and instantiable."""
    from docwen_plugin_document import DocumentPlugin

    plugin = DocumentPlugin()
    assert plugin.plugin_id == "docwen_plugin_document"


def test_docx_plugin_manifest_has_all_routes() -> None:
    """The plugin manifest must declare all document family routes (22 RouteSpec)."""
    from docwen_plugin_document import DocumentPlugin

    plugin = DocumentPlugin()
    manifest = plugin.manifest
    assert manifest.plugin_id == "docwen_plugin_document"
    assert len(manifest.routes) == 22, (
        f"Expected 22 RouteSpec (docx→md, document→md, 20 SmartConverter), got {len(manifest.routes)}"
    )
    # ROUTE-DOC-001: docx → md (primary) and document → md (category alias)
    docx_to_md = [r for r in manifest.routes if r.source_format == "docx" and r.target_format == "md"]
    assert len(docx_to_md) == 1, f"Expected exactly 1 docx→md route, got {len(docx_to_md)}"
    document_to_md = [r for r in manifest.routes if r.source_format == "document" and r.target_format == "md"]
    assert len(document_to_md) == 1, f"Expected exactly 1 document→md route, got {len(document_to_md)}"
    # SmartConverter: 20 routes covering all docx/doc/odt/rtf/wps interconversion
    smart_sources = {"docx", "doc", "odt", "rtf", "wps"}
    smart_targets = {"docx", "doc", "odt", "rtf", "wps"}
    smart_found = 0
    for r in manifest.routes:
        if r.source_format in smart_sources and r.target_format in smart_targets:
            smart_found += 1
    assert smart_found >= 20, f"Expected at least 20 SmartConverter routes, got {smart_found}"


def test_docx_plugin_can_handle_routes() -> None:
    """can_handle should match registered document-family routes."""
    from docwen_plugin_document import DocumentPlugin

    plugin = DocumentPlugin()
    # Implemented routes
    assert plugin.can_handle("document", "md") is True
    assert plugin.can_handle("docx", "md") is True
    # Only the exact public Markdown target identifier is accepted.
    assert plugin.can_handle("document", "markdown") is False
    # SmartConverter routes
    assert plugin.can_handle("odt", "docx") is True, "odt→docx should be accepted (implemented)"
    assert plugin.can_handle("wps", "rtf") is True, "wps→rtf should be accepted (implemented)"
    # Rejected routes (not in document manifest)
    assert plugin.can_handle("markdown", "docx") is False
    assert plugin.can_handle("document", "pdf") is False
    # Valid action names — only empty action for standard docx→md
    assert plugin.can_handle("document", "md", "") is True
    # Optimizer actions are now handled by dedicated optimizer plugins
    assert plugin.can_handle("docx", "md", "gongwen") is False
    assert plugin.can_handle("docx", "md", "invoice_cn") is False
    # Invalid action names — must be rejected
    assert plugin.can_handle("document", "md", "invalid_action") is False
    assert plugin.can_handle("docx", "md", "unknown") is False
    assert plugin.can_handle("docx", "md", "gongwen_invoice") is False
    assert plugin.can_handle("document", "markdown", "invalid") is False


def test_docx_plugin_has_no_stale_not_implemented_stub_docstring() -> None:
    """Document plugin routes are implemented or rejected, not advertised as stubs."""
    import inspect

    from docwen_plugin_document import plugin as plugin_module

    source = inspect.getsource(plugin_module.DocumentPlugin.can_handle)
    assert "NOT_IMPLEMENTED stubs" not in source
    assert "RouteResolver can" in source


def test_docx_to_md_manifest_declares_break_separator_options() -> None:
    """DOCX->MD route schema should expose configured break separators."""
    from docwen_plugin_document import DocumentPlugin

    plugin = DocumentPlugin()
    route = next(r for r in plugin.manifest.routes if r.source_format == "docx" and r.target_format == "md")
    properties = route.options_schema["properties"]

    assert properties["page_break_separator"]["default"] == "---"
    assert properties["section_break_separator"]["default"] == "***"
    assert properties["horizontal_rule_separator"]["default"] == "___"
    for key in ("page_break_separator", "section_break_separator", "horizontal_rule_separator"):
        assert set(properties[key]["enum"]) == {"---", "***", "___", "ignore"}
        assert properties[key]["x-docwen-status"] == "implemented"


def test_docx_to_md_manifest_ocr_placement_default_matches_export_semantics() -> None:
    """DOCX->MD schema default should match current global Markdown export semantics."""
    from docwen_plugin_document import DocumentPlugin

    plugin = DocumentPlugin()
    route = next(r for r in plugin.manifest.routes if r.source_format == "docx" and r.target_format == "md")
    properties = route.options_schema["properties"]

    assert properties["ocr_placement"]["enum"] == ["image_md", "main_md"]
    assert properties["ocr_placement"]["default"] == "main_md"


def test_docx_to_md_manifest_declares_formatting_preservation_options() -> None:
    """DOCX->MD route schema should expose consumed formatting switches."""
    from docwen_plugin_document import DocumentPlugin

    plugin = DocumentPlugin()
    route = next(r for r in plugin.manifest.routes if r.source_format == "docx" and r.target_format == "md")
    properties = route.options_schema["properties"]

    assert properties["preserve_formatting"]["default"] is True
    assert properties["preserve_heading_formatting"]["default"] is False
    assert properties["preserve_table_header_formatting"]["default"] is False
    for key in ("preserve_formatting", "preserve_heading_formatting", "preserve_table_header_formatting"):
        assert properties[key]["type"] == "boolean"
        assert properties[key]["x-docwen-status"] == "implemented"


def test_docx_to_md_manifest_declares_advanced_style_override_options() -> None:
    """DOCX->MD route schema should expose plugin-level style detector overrides."""
    from docwen_plugin_document import DocumentPlugin

    plugin = DocumentPlugin()
    route = next(r for r in plugin.manifest.routes if r.source_format == "docx" and r.target_format == "md")
    properties = route.options_schema["properties"]

    for key in ("code_block_style_aliases", "quote_style_aliases", "quote_generic_names"):
        assert properties[key]["type"] == "array"
        assert properties[key]["items"] == {"type": "string"}
        assert properties[key]["default"] == []
        assert properties[key]["x-docwen-status"] == "implemented"
        assert "Advanced DOCX→MD style override" in properties[key]["description"]


def test_docx_to_md_manifest_declares_consumed_markdown_export_options() -> None:
    """DOCX/document->MD schemas should declare every converter-consumed request option."""
    from docwen_plugin_document import DocumentPlugin

    plugin = DocumentPlugin()
    routes_by_source = {
        route.source_format: route
        for route in plugin.manifest.routes
        if route.target_format == "md" and route.source_format in {"docx", "document"}
    }

    consumed_options = {
        "to_md_keep_images",
        "remove_numbering",
        "add_numbering",
        "numbering_scheme",
        "to_md_enable_ocr",
        "ocr_language",
        "locale",
        "yaml_key_labels",
        "image_mode",
        "ocr_placement",
        "image_link_style",
        "table_merge_strategy",
        "preserve_formatting",
        "preserve_heading_formatting",
        "preserve_table_header_formatting",
        "page_break_separator",
        "section_break_separator",
        "horizontal_rule_separator",
        "code_block_style_aliases",
        "quote_style_aliases",
        "quote_generic_names",
    }

    assert set(routes_by_source) == {"docx", "document"}
    for source, route in routes_by_source.items():
        properties = route.options_schema["properties"]
        missing = consumed_options - set(properties)
        assert not missing, f"{source}->md schema misses consumed options: {sorted(missing)}"

        for key in consumed_options:
            assert properties[key]["x-docwen-status"] == "implemented"


def test_docx_plugin_only_depends_on_core() -> None:
    """Plugin must NOT import runtime, application, gui, cli, or other plugins."""
    import sys

    forbidden = {
        "docwen_runtime",
        "docwen_application",
        "docwen_gui",
        "docwen_cli",
        "docwen_bundle",
    }

    plugin_modules = {k for k in sys.modules if k.startswith("docwen_plugin_document")}

    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            if val_mod:
                for forbidden_pkg in forbidden:
                    assert not val_mod.startswith(forbidden_pkg), (
                        f"{mod_name}.{attr} depends on {val_mod} (forbidden: {forbidden_pkg})"
                    )
