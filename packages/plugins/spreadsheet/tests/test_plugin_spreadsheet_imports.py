"""Test that docwen_plugin_spreadsheet can be imported and satisfies contracts."""

import pytest

pytestmark = pytest.mark.unit

ALL_SMARTSHEET_PAIRS = [
    ("xlsx", "xls"),
    ("xlsx", "ods"),
    ("xlsx", "et"),
    ("xls", "xlsx"),
    ("ods", "xlsx"),
    ("et", "xlsx"),
    ("xls", "ods"),
    ("xls", "et"),
    ("ods", "xls"),
    ("ods", "et"),
    ("et", "xls"),
    ("et", "ods"),
    ("csv", "xls"),
    ("csv", "ods"),
    ("xls", "csv"),
    ("ods", "csv"),
    ("et", "csv"),
]


def test_plugin_spreadsheet_importable() -> None:
    """docwen_plugin_spreadsheet should be importable with version."""
    import docwen_plugin_spreadsheet

    assert docwen_plugin_spreadsheet.__version__ == "0.1.0"


def test_spreadsheet_plugin_can_be_instantiated() -> None:
    """SpreadsheetPlugin should be importable and instantiable."""
    from docwen_plugin_spreadsheet import SpreadsheetPlugin

    plugin = SpreadsheetPlugin()
    assert plugin.plugin_id == "docwen_plugin_spreadsheet"


def test_spreadsheet_plugin_manifest_has_routes() -> None:
    """The plugin manifest should declare all expected routes."""
    from docwen_plugin_spreadsheet import SpreadsheetPlugin

    plugin = SpreadsheetPlugin()
    manifest = plugin.manifest
    assert manifest.plugin_id == "docwen_plugin_spreadsheet"

    route_map = {(r.source_format, r.target_format, r.action_name): r for r in manifest.routes}

    # Core conversion
    assert ("spreadsheet", "md", "") in route_map, "Missing spreadsheet→md route"
    assert ("xlsx", "md", "") in route_map, "Missing xlsx→md route"
    assert ("csv", "md", "") in route_map, "Missing csv→md route"
    assert ("tsv", "md", "") in route_map, "Missing tsv→md route"
    assert ("csv", "xlsx", "") in route_map, "Missing csv→xlsx route"
    assert ("xlsx", "csv", "") in route_map, "Missing xlsx→csv route"
    assert ("tsv", "xlsx", "") in route_map, "Missing tsv→xlsx route"
    assert ("xlsx", "tsv", "") in route_map, "Missing xlsx→tsv route"

    # Action route
    assert ("spreadsheet", "xlsx", "merge_tables") in route_map, "Missing merge_tables action route"

    missing_smart_sheet_routes = [
        f"{source}->{target}" for source, target in ALL_SMARTSHEET_PAIRS if (source, target, "") not in route_map
    ]
    assert missing_smart_sheet_routes == []


def test_spreadsheet_plugin_manifest_uses_public_ocr_placement_values() -> None:
    """Reserved OCR placement schema must use the same public values as CLI/config."""
    from docwen_plugin_spreadsheet.manifest import SPREADSHEET_TO_MD_OPTIONS_SCHEMA

    schema = SPREADSHEET_TO_MD_OPTIONS_SCHEMA["properties"]["ocr_placement"]
    assert schema["enum"] == ["image_md", "main_md"]
    assert "text_block" not in schema["enum"]
    assert "after_table" not in schema["enum"]


def test_merge_tables_manifest_declares_only_consumed_route_options() -> None:
    """The old base-table CLI flag must not leak into the plugin schema."""
    from docwen_plugin_spreadsheet.manifest import TABLE_MERGE_OPTIONS_SCHEMA

    properties = TABLE_MERGE_OPTIONS_SCHEMA["properties"]

    assert set(properties) == {"merge_mode", "offset_range"}
    assert properties["merge_mode"]["enum"] == ["row", "col", "cell"]
    assert TABLE_MERGE_OPTIONS_SCHEMA["required"] == ["merge_mode"]
    assert "base_table" not in properties


def test_xlsx_to_ods_manifest_declares_request_only_protection_options() -> None:
    from docwen_plugin_spreadsheet import SpreadsheetPlugin

    route = next(
        route
        for route in SpreadsheetPlugin().manifest.routes
        if (route.source_format, route.target_format, route.action_name) == ("xlsx", "ods", "")
    )
    properties = route.options_schema["properties"]

    assert set(properties) == {
        "spreadsheet_password",
        "allow_spreadsheet_protection_loss",
    }
    assert properties["spreadsheet_password"]["format"] == "password"
    assert properties["spreadsheet_password"]["writeOnly"] is True
    assert properties["allow_spreadsheet_protection_loss"]["default"] is False


def test_spreadsheet_to_md_manifest_declares_consumed_markdown_export_options() -> None:
    """Spreadsheet-family -> MD routes should expose all consumed Markdown export options."""
    from docwen_plugin_spreadsheet import SpreadsheetPlugin

    plugin = SpreadsheetPlugin()
    route_map = {
        (route.source_format, route.target_format, route.action_name): route for route in plugin.manifest.routes
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
        "table_merge_strategy",
    }
    markdown_sources = {"spreadsheet", "xlsx", "csv", "tsv", "xls", "ods", "et"}

    for source in markdown_sources:
        properties = route_map[(source, "md", "")].options_schema["properties"]
        missing = consumed_options - set(properties)
        assert not missing, f"{source}->md schema misses consumed options: {sorted(missing)}"

    for key in consumed_options - {"to_md_keep_images", "table_merge_strategy"}:
        schema = route_map[("xlsx", "md", "")].options_schema["properties"][key]
        assert schema["x-docwen-status"] == "implemented"

    for source, target, action in (
        ("csv", "xlsx", ""),
        ("xlsx", "csv", ""),
        ("tsv", "xlsx", ""),
        ("xlsx", "tsv", ""),
        ("spreadsheet", "xlsx", "merge_tables"),
    ):
        properties = route_map[(source, target, action)].options_schema["properties"]
        assert not (consumed_options & set(properties))


def test_spreadsheet_plugin_can_handle_routes() -> None:
    """can_handle should match registered routes and reject invalid ones."""
    from docwen_plugin_spreadsheet import SpreadsheetPlugin

    plugin = SpreadsheetPlugin()

    # Core routes
    assert plugin.can_handle("spreadsheet", "md") is True
    assert plugin.can_handle("xlsx", "md") is True
    assert plugin.can_handle("csv", "md") is True
    assert plugin.can_handle("tsv", "md") is True
    assert plugin.can_handle("csv", "xlsx") is True
    assert plugin.can_handle("xlsx", "csv") is True
    assert plugin.can_handle("tsv", "xlsx") is True
    assert plugin.can_handle("xlsx", "tsv") is True

    # Bridge-backed routes
    assert plugin.can_handle("xls", "md") is True
    assert plugin.can_handle("ods", "md") is True
    assert plugin.can_handle("et", "md") is True

    # SmartSheetConverter routes
    for source, target in ALL_SMARTSHEET_PAIRS:
        assert plugin.can_handle(source, target) is True, f"Missing SmartSheet route {source}->{target}"

    # Action routes
    assert plugin.can_handle("spreadsheet", "xlsx", "merge_tables") is True

    # Invalid routes
    assert plugin.can_handle("markdown", "docx") is False
    assert plugin.can_handle("document", "pdf") is False
    assert plugin.can_handle("image", "md") is False
    # Cannot convert to self
    assert plugin.can_handle("xlsx", "xlsx") is False
    assert plugin.can_handle("csv", "csv") is False
    assert plugin.can_handle("xls", "xls") is False
    assert plugin.can_handle("ods", "ods") is False
    assert plugin.can_handle("et", "et") is False


def test_spreadsheet_plugin_only_depends_on_core() -> None:
    """Plugin must NOT import runtime, application, gui, cli, or other plugins."""
    import sys

    forbidden = {
        "docwen_runtime",
        "docwen_application",
        "docwen_gui",
        "docwen_cli",
        "docwen_bundle",
    }

    # Also forbid importing other plugins
    other_plugins = {
        "docwen_plugin_document",
        "docwen_plugin_presentation",
        "docwen_plugin_markup",
        "docwen_plugin_markdown",
        "docwen_plugin_image",
        "docwen_plugin_layout",
        "docwen_plugin_print",
        "docwen_plugin_optimizer_gongwen",
        "docwen_plugin_optimizer_invoice_cn",
        "docwen_plugin_proofread",
        "docwen_plugin_template",
    }

    plugin_modules = {k for k in sys.modules if k.startswith("docwen_plugin_spreadsheet")}

    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            if val_mod:
                for forbidden_pkg in forbidden | other_plugins:
                    if val_mod.startswith(forbidden_pkg):
                        # Allow imports of our own package
                        if val_mod.startswith("docwen_plugin_spreadsheet"):
                            continue
                        raise AssertionError(f"{mod_name}.{attr} depends on {val_mod} (forbidden: {forbidden_pkg})")
