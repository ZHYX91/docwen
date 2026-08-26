"""Verify markdown plugin can be imported and exposes expected public API."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.contract


def test_plugin_can_be_imported():
    """Smoke test: plugin package is importable."""
    import docwen_plugin_markdown

    assert docwen_plugin_markdown is not None


def test_markdown_plugin_class_exists():
    """MarkdownPlugin class is directly importable."""
    from docwen_plugin_markdown import MarkdownPlugin

    assert MarkdownPlugin is not None


def test_plugin_instantiable():
    """MarkdownPlugin can be instantiated."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    assert plugin.plugin_id == "docwen_plugin_markdown"


def test_manifest_has_routes():
    """The manifest declares at least 9 routes."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    manifest = plugin.manifest
    assert len(manifest.routes) >= 9


def test_manifest_id_and_version():
    """Manifest has correct plugin_id and version fields."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    manifest = plugin.manifest
    assert manifest.plugin_id == "docwen_plugin_markdown"
    assert manifest.version == "0.1.0"


@pytest.mark.parametrize(
    "source,target,action,expected",
    [
        ("markdown", "docx", "", True),
        ("md", "docx", "", False),
        ("markdown", "doc", "", True),
        ("markdown", "odt", "", True),
        ("markdown", "rtf", "", True),
        ("markdown", "wps", "", True),
        ("markdown", "pdf", "", True),
        ("markdown", "xlsx", "", True),
        ("markdown", "xls", "", True),
        ("markdown", "ods", "", True),
        ("markdown", "csv", "", True),
        ("markdown", "md", "process_md_numbering", True),
        # Negative cases
        ("pdf", "docx", "", False),
        ("docx", "md", "", False),
        ("markdown", "png", "", False),
    ],
)
def test_can_handle(source, target, action, expected):
    """can_handle correctly matches declared routes."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    assert plugin.can_handle(source, target, action) == expected


def test_manifest_serialization():
    """Manifest can be serialized to/from dict."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    manifest = plugin.manifest
    d = manifest.to_dict()
    assert d["plugin_id"] == "docwen_plugin_markdown"
    assert len(d["routes"]) >= 9


def test_manifest_template_option_matches_document_and_spreadsheet_template_targets():
    """Keep the manifest honest about which Markdown routes consume templates."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    routes_by_target = {route.target_format: route for route in plugin.manifest.routes}

    for target in {"docx", "doc", "odt", "rtf", "wps", "pdf"}:
        properties = routes_by_target[target].options_schema.get("properties", {})
        assert "template_name" in properties

    for target in {"xlsx", "xls", "ods", "csv"}:
        properties = routes_by_target[target].options_schema.get("properties", {})
        assert "template_name" in properties


def test_manifest_limits_word_native_render_mode_to_md_to_document_routes() -> None:
    """word_native is a DOCX-intermediate output option, not a Markdown numbering action option."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    routes_by_action = {route.action_name: route for route in plugin.manifest.routes}
    routes_by_target = {route.target_format: route for route in plugin.manifest.routes}

    for target in {"docx", "doc", "odt", "rtf", "wps", "pdf"}:
        properties = routes_by_target[target].options_schema.get("properties", {})
        assert properties["heading_numbering_render_mode"]["enum"] == ["text", "word_native"]

    numbering_properties = routes_by_action["process_md_numbering"].options_schema.get("properties", {})
    assert "heading_numbering_render_mode" not in numbering_properties


def test_manifest_declares_docx_intermediate_table_header_formatting_option() -> None:
    """table_header_formatting_mode is consumed by MD->DOCX rendering routes only."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    routes_by_action = {route.action_name: route for route in plugin.manifest.routes}
    routes_by_target = {route.target_format: route for route in plugin.manifest.routes}

    for target in {"docx", "doc", "odt", "rtf", "wps", "pdf"}:
        properties = routes_by_target[target].options_schema.get("properties", {})
        assert properties["table_header_formatting_mode"]["enum"] == ["apply", "keep", "remove"]
        assert properties["table_header_formatting_mode"]["default"] == "remove"

    for target in {"xlsx", "xls", "ods", "csv"}:
        properties = routes_by_target[target].options_schema.get("properties", {})
        assert "table_header_formatting_mode" not in properties

    numbering_properties = routes_by_action["process_md_numbering"].options_schema.get("properties", {})
    assert "table_header_formatting_mode" not in numbering_properties


def test_manifest_declares_consumed_docx_intermediate_rendering_options() -> None:
    """MD->DOCX rendering options consumed by the converter must be route-declared."""
    from docwen_plugin_markdown import MarkdownPlugin

    plugin = MarkdownPlugin()
    routes_by_action = {route.action_name: route for route in plugin.manifest.routes}
    routes_by_target = {route.target_format: route for route in plugin.manifest.routes}

    consumed_docx_options = {
        "locale",
        "remove_numbering",
        "add_numbering",
        "numbering_scheme",
        "heading_numbering_render_mode",
        "formatting_mode",
        "heading_formatting_mode",
        "table_header_formatting_mode",
        "code_font",
        "code_background_color",
        "table_style_mode",
        "builtin_style_key",
        "custom_style_name",
        "heading_merge_mode",
        "heading_merge_punctuation",
        "template_name",
        "hr_mapping",
    }
    spreadsheet_template_options = {"template_name"}
    numbering_action_options = {
        "remove_numbering",
        "add_numbering",
        "numbering_scheme",
    }
    docx_render_only_options = consumed_docx_options - spreadsheet_template_options - numbering_action_options

    for target in {"docx", "doc", "odt", "rtf", "wps", "pdf"}:
        properties = routes_by_target[target].options_schema.get("properties", {})
        assert consumed_docx_options <= set(properties)

    for target in {"xlsx", "xls", "ods", "csv"}:
        properties = routes_by_target[target].options_schema.get("properties", {})
        assert set(properties) == spreadsheet_template_options

    numbering_properties = routes_by_action["process_md_numbering"].options_schema.get("properties", {})
    assert set(numbering_properties) == numbering_action_options
    assert not docx_render_only_options & set(numbering_properties)


def test_plugin_only_imports_allowed_deps():
    """Plugin exports must not depend on runtime/apps/other plugins."""
    import importlib
    import sys

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
        "docwen_plugin_layout",
        "docwen_plugin_print",
        "docwen_plugin_optimizer_gongwen",
        "docwen_plugin_optimizer_invoice_cn",
        "docwen_plugin_proofread",
    }
    before = set(sys.modules)
    importlib.import_module("docwen_plugin_markdown")
    introduced = set(sys.modules) - before
    assert not {name for name in introduced if any(name.startswith(pkg) for pkg in forbidden)}

    plugin_modules = {name for name in sys.modules if name.startswith("docwen_plugin_markdown")}
    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            if not val_mod:
                continue
            for forbidden_pkg in forbidden:
                assert not val_mod.startswith(forbidden_pkg), f"{mod_name}.{attr} depends on {val_mod}"


def test_markdown_plugin_source_does_not_depend_on_optimizer_plugins() -> None:
    """Markdown plugin production code owns MD->DOCX field handling locally."""
    from pathlib import Path

    plugin_root = Path(__file__).resolve().parents[1] / "src" / "docwen_plugin_markdown"
    production_source = "\n".join(path.read_text(encoding="utf-8") for path in sorted(plugin_root.rglob("*.py")))

    assert "docwen_plugin_optimizer_gongwen" not in production_source
    assert "docwen_plugin_optimizer_invoice_cn" not in production_source


def test_markdown_plugin_has_no_stale_not_implemented_stub_docstring() -> None:
    """Markdown plugin manifest routes are implemented or bridge-backed."""
    import inspect

    from docwen_plugin_markdown import plugin as plugin_module

    source = inspect.getsource(plugin_module.MarkdownPlugin.can_handle)
    assert "NOT_IMPLEMENTED stubs" not in source
    assert "RouteResolver can" in source


def test_common_utils_has_no_dead_docx_ast_alias() -> None:
    from docwen_plugin_markdown import common_utils

    assert not hasattr(common_utils, "BUILD_DOCX_FROM_AST")


def test_yaml_processor_only_owns_front_matter_extraction() -> None:
    """YAML field processing lives in field_registry, not legacy stubs."""
    import inspect

    from docwen_plugin_markdown import yaml_processor

    assert hasattr(yaml_processor, "extract_yaml_front_matter")
    assert not hasattr(yaml_processor, "process_yaml_fields")
    assert not hasattr(yaml_processor, "render_yaml_fields_to_docx")

    source = inspect.getsource(yaml_processor)
    assert "Phase 1: no-op stub" not in source
    assert "Phase 2 implements" not in source


def test_markdown_pipeline_modules_have_no_phase_stub_residue() -> None:
    """Markdown pipeline modules must not carry unused migration-era stubs."""
    import inspect

    from docwen_plugin_markdown import (
        ast_transforms,
        common_utils,
        mistune_extensions,
        preprocessor,
        renderer,
        renderer_inlines,
        renderer_utils,
        template_filler,
    )

    assert not hasattr(preprocessor, "_TAG_MAP")

    source = (
        inspect.getsource(ast_transforms)
        + inspect.getsource(common_utils)
        + inspect.getsource(mistune_extensions)
        + inspect.getsource(preprocessor)
        + inspect.getsource(renderer)
        + inspect.getsource(renderer_inlines)
        + inspect.getsource(renderer_utils)
        + inspect.getsource(template_filler)
    )
    assert "Phase 2 stub" not in source
    assert "Phase 2" not in source
    assert "Phase 3" not in source
    assert "_common.py" not in source
    assert "formerly in" not in source
    assert "will be handled soon" not in source
    assert "will be refactored into domain-specific modules" not in source


def test_word_numbering_production_code_uses_core_adapter() -> None:
    """Word-numbering production code has one canonical implementation."""
    from pathlib import Path

    project_root = Path(__file__).resolve().parents[4]
    converter_source = (
        project_root
        / "packages"
        / "plugins"
        / "markdown"
        / "src"
        / "docwen_plugin_markdown"
        / "to_docx"
        / "converter.py"
    ).read_text(encoding="utf-8")
    heading_source = (
        project_root
        / "packages"
        / "plugins"
        / "markdown"
        / "src"
        / "docwen_plugin_markdown"
        / "to_docx"
        / "heading_numbering.py"
    ).read_text(encoding="utf-8")
    production_source = converter_source + heading_source
    assert "docwen_core.text.numbering_word_adapter" in production_source
    assert "docwen_plugin_markdown.to_docx.numbering_translator" not in production_source
    assert not (
        project_root
        / "packages"
        / "plugins"
        / "markdown"
        / "src"
        / "docwen_plugin_markdown"
        / "to_docx"
        / "numbering_translator.py"
    ).exists()
