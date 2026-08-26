"""Repository guards for plugin option consumption and reserved route scope."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]

_OPTION_GET_LITERAL_RE = re.compile(r'\b(?:options|opts|context\.request\.options)\.get\("([^"]+)"')
_CONTEXT_REQUEST_OPTIONS_GET_RE = re.compile(r'getattr\(context\.request,\s*"options",\s*\{\}\)\.get\("([^"]+)"')
_REQUEST_OPTIONS_INDEX_RE = re.compile(r'"([^"]+)"\s+in\s+request_options|request_options\["([^"]+)"\]')
_REQ_OPTS_INDEX_RE = re.compile(r'"([^"]+)"\s+in\s+req_opts|req_opts\["([^"]+)"\]')
_ENGINE_GET_LITERAL_RE = re.compile(r'engine\.get\("([^"]+)"')
_DYNAMIC_OPTION_HELPER_RE = re.compile(
    r'(?:_option_or_config|_option_value|_option_nonblank|_option_strings)\(\s*options,\s*"([^"]+)"',
    re.DOTALL,
)


def _plugin_request_option_read_keys() -> set[str]:
    source_root = PROJECT_ROOT / "packages" / "plugins"
    keys: set[str] = set()
    patterns = (
        _OPTION_GET_LITERAL_RE,
        _CONTEXT_REQUEST_OPTIONS_GET_RE,
        _REQUEST_OPTIONS_INDEX_RE,
        _REQ_OPTS_INDEX_RE,
        _ENGINE_GET_LITERAL_RE,
        _DYNAMIC_OPTION_HELPER_RE,
    )
    for path in source_root.glob("**/src/**/*.py"):
        source = path.read_text(encoding="utf-8")
        for pattern in patterns:
            for match in pattern.finditer(source):
                key = next(group for group in match.groups() if group)
                keys.add(key)
    return keys


def _plugin_manifest_option_keys() -> set[str]:
    keys, _reserved_keys = _plugin_manifest_option_key_sets()
    return keys


def _plugin_manifests() -> dict[str, Any]:
    from docwen_plugin_document import PLUGIN_MANIFEST as DOCUMENT_MANIFEST
    from docwen_plugin_image import PLUGIN_MANIFEST as IMAGE_MANIFEST
    from docwen_plugin_layout import PLUGIN_MANIFEST as LAYOUT_MANIFEST
    from docwen_plugin_markdown import PLUGIN_MANIFEST as MARKDOWN_MANIFEST
    from docwen_plugin_markup import PLUGIN_MANIFEST as MARKUP_MANIFEST
    from docwen_plugin_optimizer_gongwen import PLUGIN_MANIFEST as GONGWEN_MANIFEST
    from docwen_plugin_optimizer_invoice_cn import PLUGIN_MANIFEST as INVOICE_MANIFEST
    from docwen_plugin_presentation import PLUGIN_MANIFEST as PRESENTATION_MANIFEST
    from docwen_plugin_print import PLUGIN_MANIFEST as PRINT_MANIFEST
    from docwen_plugin_proofread import PLUGIN_MANIFEST as PROOFREAD_MANIFEST
    from docwen_plugin_spreadsheet import PLUGIN_MANIFEST as SPREADSHEET_MANIFEST

    return {
        "docwen_plugin_document": DOCUMENT_MANIFEST,
        "docwen_plugin_image": IMAGE_MANIFEST,
        "docwen_plugin_layout": LAYOUT_MANIFEST,
        "docwen_plugin_markdown": MARKDOWN_MANIFEST,
        "docwen_plugin_markup": MARKUP_MANIFEST,
        "docwen_plugin_optimizer_gongwen": GONGWEN_MANIFEST,
        "docwen_plugin_optimizer_invoice_cn": INVOICE_MANIFEST,
        "docwen_plugin_presentation": PRESENTATION_MANIFEST,
        "docwen_plugin_print": PRINT_MANIFEST,
        "docwen_plugin_proofread": PROOFREAD_MANIFEST,
        "docwen_plugin_spreadsheet": SPREADSHEET_MANIFEST,
    }


def _plugin_manifest_option_key_sets() -> tuple[set[str], set[str]]:
    keys: set[str] = set()
    reserved_keys: set[str] = set()
    for manifest in _plugin_manifests().values():
        for route in manifest.routes:
            for key, schema in ((route.options_schema or {}).get("properties", {}) or {}).items():
                keys.add(key)
                if isinstance(schema, dict) and schema.get("x-docwen-status") == "reserved":
                    reserved_keys.add(key)
    return keys, reserved_keys


def test_plugin_request_option_reads_are_manifest_declared_or_internal() -> None:
    """Every literal plugin request-option read should be schema-backed or documented internal."""
    request_read_keys = _plugin_request_option_read_keys()
    manifest_keys = _plugin_manifest_option_keys()
    documented_internal_keys = {
        "output_dir",  # Gongwen image extraction helper output; not a public route option.
    }

    assert "skip_code_blocks" in request_read_keys
    assert "enable_symbol_pairing" in request_read_keys
    assert "table_header_formatting_mode" in request_read_keys
    assert "heading_merge_mode" in request_read_keys

    undeclared = sorted(request_read_keys - manifest_keys - documented_internal_keys)
    assert undeclared == []


def test_plugin_manifest_options_are_consumed() -> None:
    """Route schemas must not expose options without consumption evidence."""
    request_read_keys = _plugin_request_option_read_keys()
    manifest_keys, reserved_keys = _plugin_manifest_option_key_sets()

    assert reserved_keys == set()
    assert "template_name" in request_read_keys
    assert "skip_quote_blocks" in request_read_keys

    unconsumed_public_keys = sorted(manifest_keys - request_read_keys - reserved_keys)
    assert unconsumed_public_keys == []


def test_image_reused_option_keys_stay_route_scoped_and_consumed() -> None:
    """Image routes expose only the options consumed by that route."""
    from docwen_plugin_image.manifest import build_manifest

    route_map = {
        (route.source_format, route.target_format, route.action_name): route for route in build_manifest().routes
    }
    image_common_source = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "src" / "docwen_plugin_image" / "_common.py"
    ).read_text(encoding="utf-8")
    image_pdf_source = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "src" / "docwen_plugin_image" / "to_pdf" / "converter.py"
    ).read_text(encoding="utf-8")
    image_merge_source = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "src" / "docwen_plugin_image" / "merge" / "converter.py"
    ).read_text(encoding="utf-8")

    generic_image_props = route_map[("image", "image", "")].options_schema["properties"]
    explicit_image_props = route_map[("image", "webp", "")].options_schema["properties"]
    image_pdf_props = route_map[("image", "pdf", "")].options_schema["properties"]
    merge_tiff_props = route_map[("image", "tif", "merge_images_to_tiff")].options_schema["properties"]

    for props in (generic_image_props, explicit_image_props):
        assert props["compress_mode"].get("x-docwen-status") != "reserved"
        assert props["size_limit"].get("x-docwen-status") != "reserved"
        assert props["size_unit"].get("x-docwen-status") != "reserved"
    assert 'options.get("compress_mode")' in image_common_source
    assert 'options.get("size_limit")' in image_common_source
    assert 'options.get("size_unit", "KB")' in image_common_source
    assert "tiff_mode" not in generic_image_props
    assert "tiff_mode" not in explicit_image_props

    assert image_pdf_props["quality_mode"].get("x-docwen-status") != "reserved"
    assert set(image_pdf_props) == {"quality_mode"}
    assert 'options.get("quality_mode")' in image_pdf_source
    assert 'options.get("pdf_quality")' not in image_pdf_source
    assert 'options.get("compress_mode")' not in image_pdf_source
    assert 'options.get("size_limit")' not in image_pdf_source

    assert merge_tiff_props["mode"].get("x-docwen-status") != "reserved"
    assert merge_tiff_props["keep_alpha"].get("x-docwen-status") != "reserved"
    assert set(merge_tiff_props) == {"mode", "keep_alpha"}
    assert 'options.get("mode")' in image_merge_source
    assert 'options.get("keep_alpha")' in image_merge_source
    assert 'options.get("compress")' not in image_merge_source
    assert 'options.get("size_limit")' not in image_merge_source
    assert 'options.get("quality_mode")' not in image_merge_source


def test_route_manifests_expose_no_reserved_options() -> None:
    """Incomplete options belong in private plans, not the public route contract."""
    reserved_entries: set[tuple[str, str, str, str, str]] = set()
    for plugin_id, manifest in _plugin_manifests().items():
        for route in manifest.routes:
            props = (route.options_schema or {}).get("properties", {}) or {}
            for key, schema in props.items():
                if isinstance(schema, dict) and schema.get("x-docwen-status") == "reserved":
                    reserved_entries.add(
                        (
                            plugin_id,
                            route.source_format,
                            route.target_format,
                            route.action_name,
                            key,
                        )
                    )

    assert reserved_entries == set()


def test_every_plugin_route_option_schema_is_closed() -> None:
    """Unknown route options must fail closed instead of becoming compatibility shims."""
    open_routes: list[str] = []
    for plugin_id, manifest in _plugin_manifests().items():
        for route in manifest.routes:
            if route.options_schema.get("additionalProperties") is not False:
                open_routes.append(f"{plugin_id}:{route.source_format}->{route.target_format}:{route.action_name}")

    assert open_routes == []
