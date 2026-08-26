"""Invoice plugin import and manifest contract tests."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_plugin_invoice_importable() -> None:
    import docwen_plugin_optimizer_invoice_cn

    assert docwen_plugin_optimizer_invoice_cn.__version__ == "0.1.0"


def test_invoice_plugin_can_be_instantiated() -> None:
    from docwen_plugin_optimizer_invoice_cn import InvoicePlugin

    plugin = InvoicePlugin()
    assert plugin.plugin_id == "docwen_plugin_optimizer_invoice_cn"


def test_invoice_plugin_manifest_declares_routes() -> None:
    from docwen_plugin_optimizer_invoice_cn import InvoicePlugin

    manifest = InvoicePlugin().manifest
    assert manifest.plugin_id == "docwen_plugin_optimizer_invoice_cn"
    routes = {(r.source_format, r.target_format, r.action_name) for r in manifest.routes}

    # Action routes for invoice_cn
    assert ("pdf", "md", "invoice_cn") in routes
    assert ("ofd", "md", "invoice_cn") in routes
    assert ("image", "md", "invoice_cn") in routes

    assert manifest.requires == []
    assert [resource.to_dict() for resource in manifest.optimization_resources] == [
        {
            "id": "invoice_cn",
            "name": "Chinese invoice optimization",
            "action_name": "invoice_cn",
        }
    ]
    assert "supported_actions" not in manifest.extra

    # Verify plugin-specific extra metadata.
    assert "invoice_cn_yaml_schema" in manifest.extra
    assert len(manifest.extra["invoice_cn_yaml_schema"]) == 20  # 20 fields


def test_invoice_plugin_can_handle_manifest_routes() -> None:
    from docwen_plugin_optimizer_invoice_cn import InvoicePlugin

    plugin = InvoicePlugin()

    # Action routes
    assert plugin.can_handle("pdf", "md", "invoice_cn") is True
    assert plugin.can_handle("ofd", "md", "invoice_cn") is True
    assert plugin.can_handle("image", "md", "invoice_cn") is True

    # Negative cases — need both action and format match
    assert plugin.can_handle("pdf", "md") is False  # no action
    assert plugin.can_handle("ofd", "md") is False
    assert plugin.can_handle("image", "md") is False
    assert plugin.can_handle("pdf", "pdf", "invoice_cn") is False  # wrong target
    assert plugin.can_handle("docx", "md", "invoice_cn") is False  # wrong source


def test_ocr_option_is_available() -> None:
    """The to_md_enable_ocr option should be available (no longer reserved)."""
    from docwen_plugin_optimizer_invoice_cn import InvoicePlugin

    manifest = InvoicePlugin().manifest
    assert "reserved_features" not in manifest.extra
    for route in manifest.routes:
        schema = route.options_schema
        props = schema.get("properties", {})
        ocr_prop = props.get("to_md_enable_ocr", {})
        assert ocr_prop.get("default") is False, f"to_md_enable_ocr default should be False in route {route.label}"
        assert "type" in ocr_prop
        assert ocr_prop["type"] == "boolean"

        language_prop = props.get("ocr_language", {})
        assert language_prop.get("default") == "auto", f"ocr_language default should be auto in route {route.label}"
        assert language_prop.get("enum") == [
            "auto",
            "chinese",
            "chinese_cht",
            "english",
            "japanese",
            "korean",
            "latin",
            "cyrillic",
        ]


def test_invoice_cn_manifest_declares_consumed_action_options() -> None:
    """Invoice action routes should expose only the options the converter consumes."""
    from docwen_plugin_optimizer_invoice_cn import InvoicePlugin

    route_map = {
        (route.source_format, route.target_format, route.action_name): route
        for route in InvoicePlugin().manifest.routes
    }
    consumed_options = {"to_md_enable_ocr", "ocr_language", "locale", "yaml_key_labels"}

    assert set(route_map) == {
        ("pdf", "md", "invoice_cn"),
        ("ofd", "md", "invoice_cn"),
        ("image", "md", "invoice_cn"),
    }
    for route in route_map.values():
        props = route.options_schema["properties"]
        assert set(props) == consumed_options
        assert props["ocr_language"]["enum"] == [
            "auto",
            "chinese",
            "chinese_cht",
            "english",
            "japanese",
            "korean",
            "latin",
            "cyrillic",
        ]
        assert props["locale"]["x-docwen-status"] == "implemented"
        assert props["yaml_key_labels"]["x-docwen-status"] == "implemented"


def test_invoice_ocr_docs_do_not_describe_image_route_as_reserved() -> None:
    """Image invoice OCR is wired; docs must not describe it as a reserved stub."""
    project_root = Path(__file__).resolve().parents[5]
    invoice_root = project_root / "packages" / "plugins" / "optimizers" / "invoice_cn"
    source_text = "\n".join(
        (invoice_root / relative_path).read_text(encoding="utf-8")
        for relative_path in (
            "src/docwen_plugin_optimizer_invoice_cn/__init__.py",
            "src/docwen_plugin_optimizer_invoice_cn/plugin.py",
            "tests/_invoice_conversions_support.py",
        )
    )

    assert 'image → md with action_name="invoice_cn" (OCR-backed' in source_text
    assert "OCR parsing for image-based invoices" in source_text
    assert "OCR reserved" not in source_text
    assert "returns NOT_IMPLEMENTED" not in source_text
    assert "declared as reserved" not in source_text


def test_invoice_plugin_only_depends_on_core() -> None:
    """Verify the invoice plugin does not import forbidden packages."""
    import sys

    import docwen_plugin_optimizer_invoice_cn  # type: ignore[unused-import]  # noqa: F401

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
        "docwen_plugin_markdown",
        "docwen_plugin_layout",
        "docwen_plugin_print",
        "docwen_plugin_proofread",
        "docwen.",
    }
    plugin_modules = {name for name in sys.modules if name.startswith("docwen_plugin_optimizer_invoice_cn")}
    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            for forbidden_pkg in forbidden:
                assert not val_mod.startswith(forbidden_pkg), f"{mod_name}.{attr} depends on {val_mod}"
