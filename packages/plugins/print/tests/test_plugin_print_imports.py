"""Print plugin import and manifest contract tests."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_plugin_print_importable() -> None:
    import docwen_plugin_print

    assert docwen_plugin_print.__version__ == "0.1.0"


def test_print_plugin_can_be_instantiated() -> None:
    from docwen_plugin_print import PrintPlugin

    plugin = PrintPlugin()
    assert plugin.plugin_id == "docwen_plugin_print"


def test_print_plugin_manifest_declares_routes() -> None:
    from docwen_plugin_print import PrintPlugin

    manifest = PrintPlugin().manifest
    assert manifest.plugin_id == "docwen_plugin_print"
    routes = {(r.source_format, r.target_format, r.action_name) for r in manifest.routes}

    # ── Document family → PDF (Office bridge) ─────────────────────────
    for src in ("docx", "doc", "odt", "rtf", "wps", "document"):
        assert (src, "pdf", "") in routes, f"missing: {src}→pdf"

    # ── OFD export is unavailable and therefore not executable ───────
    for src in ("docx", "doc", "odt", "rtf", "wps", "document"):
        assert (src, "ofd", "") not in routes

    # ── Spreadsheet family → PDF (Office bridge) ──────────────────────
    for src in ("xlsx", "xls", "ods", "et", "csv", "spreadsheet"):
        assert (src, "pdf", "") in routes, f"missing: {src}→pdf"

    # ── tsv is not registered (not implemented) ───────────────────────
    assert ("tsv", "pdf", "") not in routes

    # ── Total: 6 doc sources + 6 sheet sources = 12 ──────────────────
    assert len(manifest.routes) == 12, f"Expected 12 routes, got {len(manifest.routes)}"

    assert manifest.requires == []


def test_print_plugin_manifest_reports_ofd_as_unavailable() -> None:
    """Capability metadata keeps OFD outside the executable route surface."""
    from docwen_plugin_print import PrintPlugin

    manifest = PrintPlugin().manifest

    assert {route.target_format for route in manifest.routes} == {"pdf"}
    assert manifest.extra["unavailable_target_formats"] == ["ofd"]
    assert "not_implemented_routes" not in manifest.extra


def test_print_manifest_has_no_executable_not_implemented_ofd_route() -> None:
    """Document->OFD is capability metadata, not an executable error stub."""
    source = Path(__file__).parents[1].joinpath("src", "docwen_plugin_print", "manifest.py").read_text(encoding="utf-8")

    assert "NOT_IMPLEMENTED stub" not in source
    assert "clean not_implemented route" not in source
    assert '"unavailable_routes"' in source


def test_print_plugin_can_handle_explicit_routes() -> None:
    """can_handle() must return True for all explicitly registered routes."""
    from docwen_plugin_print import PrintPlugin

    plugin = PrintPlugin()

    # Document family → PDF
    for src in ("docx", "doc", "odt", "rtf", "wps", "document"):
        assert plugin.can_handle(src, "pdf") is True, f"can_handle({src}, pdf) failed"

    # Document family → OFD is deliberately not executable.
    for src in ("docx", "doc", "odt", "rtf", "wps", "document"):
        assert plugin.can_handle(src, "ofd") is False

    # Spreadsheet family → PDF
    for src in ("xlsx", "xls", "ods", "et", "csv", "spreadsheet"):
        assert plugin.can_handle(src, "pdf") is True, f"can_handle({src}, pdf) failed"

    # ── tsv is not registered ─────────────────────────────────────────
    assert plugin.can_handle("tsv", "pdf") is False

    # ── Negative cases ────────────────────────────────────────────────
    assert plugin.can_handle("pdf", "md") is False
    assert plugin.can_handle("image", "png") is False


def test_print_plugin_only_depends_on_core() -> None:
    import sys

    import docwen_plugin_print  # type: ignore[unused-import]  # noqa: F401

    forbidden = {
        "docwen_runtime",
        "docwen_application",
        "docwen_gui",
        "docwen_cli",
        "docwen_bundle",
        "docwen_plugin_document",
        "docwen_plugin_presentation",
        "docwen_plugin_markup",
        "docwen_plugin_layout",
        "docwen_plugin_spreadsheet",
        "docwen_plugin_image",
        "docwen_plugin_markdown",
        "docwen_plugin_proofread",
        "docwen.",
    }
    plugin_modules = {name for name in sys.modules if name.startswith("docwen_plugin_print")}
    for mod_name in plugin_modules:
        mod = sys.modules[mod_name]
        for attr in dir(mod):
            val = getattr(mod, attr, None)
            if val is None:
                continue
            val_mod = getattr(val, "__module__", "")
            for forbidden_pkg in forbidden:
                assert not val_mod.startswith(forbidden_pkg), f"{mod_name}.{attr} depends on {val_mod}"
