"""Gate test: manifest route-honesty consistency.

Locks the manifest == code contract for the honesty fields
``not_implemented_routes``, ``unavailable_routes`` and ``office_bridge_routes``. These are stored as
``list[HonestyRoute]`` (structured data) in each manifest's ``extra`` dict, so
consistency checks iterate structured ``(source, targets)`` fields directly —
no string parsing.

Invariants enforced:
  1. A route must not be declared as both office-bridged (implemented) and
     not-implemented — that was the exact stale-declaration bug W2 fixed
     (markdown/layout office-bridge routes were wrongly listed as
     not_implemented).
  2. Every ``(source, target)`` mentioned in either honesty list must
     correspond to a real route in the plugin's ``ALL_ROUTES`` (no phantom
     declarations of routes that do not exist).
  3. Every honesty entry is a ``HonestyRoute`` instance (guards against a
     regression to free-form strings).

These are structural invariants. Runtime unavailable-route behaviour is
covered by the per-plugin route tests (for example, ``test_print_routes.py``).
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_core.docx_styles import SHIPPED_STYLE_LOCALES
from docwen_core.models.manifest import HonestyRoute

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]


# Plugins that carry honesty metadata. Each entry is the package providing
# ``PLUGIN_MANIFEST``. Other plugins intentionally have no
# not_implemented_routes / office_bridge_routes (verified separately).
_HONESTY_PLUGINS = (
    "docwen_plugin_layout",
    "docwen_plugin_markdown",
    "docwen_plugin_print",
)


def _route_pairs(manifest) -> set[tuple[str, str]]:
    """All (source_format, target_format) pairs declared in the manifest routes."""
    return {(r.source_format, r.target_format) for r in manifest.routes}


def _honesty_pairs(entries: list[HonestyRoute]) -> set[tuple[str, str]]:
    """Expand a list of HonestyRoute into the set of (source, target) pairs it covers."""
    return {(e.source, t) for e in entries for t in e.targets}


def _load_manifest(pkg: str):
    return __import__(pkg).PLUGIN_MANIFEST


@pytest.mark.parametrize("pkg", _HONESTY_PLUGINS)
def test_manifest_honesty_fields_are_structured_when_present(pkg: str) -> None:
    """Honesty fields are optional and contain only structured declarations."""
    extra = _load_manifest(pkg).extra
    for field in ("not_implemented_routes", "office_bridge_routes", "unavailable_routes"):
        entries = extra.get(field)
        if entries is not None:
            assert isinstance(entries, list)
            assert all(isinstance(e, HonestyRoute) for e in entries), (
                f"{pkg} {field} must be list[HonestyRoute], not free-form strings"
            )


def test_office_bridge_routes_disjoint_from_not_implemented() -> None:
    """No route may be declared both office-bridged (implemented) and not-implemented."""
    for pkg in _HONESTY_PLUGINS:
        extra = _load_manifest(pkg).extra
        not_impl_pairs = _honesty_pairs(extra.get("not_implemented_routes", []) or [])
        bridge_pairs = _honesty_pairs(extra.get("office_bridge_routes", []) or [])

        overlap = not_impl_pairs & bridge_pairs
        assert not overlap, (
            f"{pkg}: routes declared both office-bridged and not-implemented (stale declaration): {sorted(overlap)}"
        )


def test_honesty_routes_exist_in_route_table() -> None:
    """Every (source, target) in honesty lists must be a real declared route."""
    for pkg in _HONESTY_PLUGINS:
        manifest = _load_manifest(pkg)
        real = _route_pairs(manifest)
        extra = manifest.extra
        not_impl_pairs = _honesty_pairs(extra.get("not_implemented_routes", []) or [])
        bridge_pairs = _honesty_pairs(extra.get("office_bridge_routes", []) or [])

        missing = sorted((not_impl_pairs | bridge_pairs) - real)
        assert not missing, f"{pkg}: honesty lists reference routes not in ALL_ROUTES: {missing}"


def test_unavailable_routes_are_absent_from_executable_route_table() -> None:
    """An unavailable capability is discovery metadata, never a dispatch route."""
    for pkg in _HONESTY_PLUGINS:
        manifest = _load_manifest(pkg)
        unavailable = _honesty_pairs(manifest.extra.get("unavailable_routes", []) or [])
        assert not (unavailable & _route_pairs(manifest)), (
            f"{pkg}: unavailable capabilities leaked into executable routes: "
            f"{sorted(unavailable & _route_pairs(manifest))}"
        )


def test_implemented_office_bridge_routes_use_concrete_layout_sources() -> None:
    """Concrete fixed-layout document routes must be honestly office-backed.

    The workflow category ``layout`` is not an admitted file format and must not
    reappear as a phantom route or honesty entry.
    """
    from docwen_plugin_layout import PLUGIN_MANIFEST

    extra = PLUGIN_MANIFEST.extra
    not_impl_pairs = _honesty_pairs(extra.get("not_implemented_routes", []))
    bridge_pairs = _honesty_pairs(extra["office_bridge_routes"])

    expected_bridged = {
        (source, target) for source in ("pdf", "ofd", "xps") for target in ("docx", "doc", "odt", "rtf")
    }
    assert expected_bridged == bridge_pairs
    assert all(source != "layout" for source, _target in bridge_pairs)
    assert not (expected_bridged & not_impl_pairs), (
        f"fixed-layout office-bridge routes wrongly marked not-implemented: {expected_bridged & not_impl_pairs}"
    )


def test_markdown_office_bridge_routes_implemented_not_not_implemented() -> None:
    """Markdown's office-bridge routes (md→doc/odt/rtf/xls/ods) must not be not-implemented."""
    from docwen_plugin_markdown import PLUGIN_MANIFEST

    extra = PLUGIN_MANIFEST.extra
    assert "not_implemented_routes" not in extra
    bridge_pairs = _honesty_pairs(extra["office_bridge_routes"])
    expected_bridged = {("markdown", t) for t in ("doc", "odt", "rtf", "xls", "ods")}
    assert expected_bridged <= bridge_pairs


def test_print_document_to_ofd_is_unavailable_not_executable() -> None:
    """Print reports OFD as unavailable without registering a dead route."""
    from docwen_plugin_print import PLUGIN_MANIFEST

    extra = PLUGIN_MANIFEST.extra
    assert "not_implemented_routes" not in extra
    unavailable_pairs = _honesty_pairs(extra["unavailable_routes"])
    assert unavailable_pairs == {("document", "ofd")}
    assert ("document", "ofd") not in _route_pairs(PLUGIN_MANIFEST)
    assert extra.get("office_bridge_routes") is None


def test_plugins_without_honesty_lists_have_none() -> None:
    """Plugins that do not declare honesty lists must not fabricate them.

    Per the W2 audit, document/presentation/spreadsheet/markup/image/gongwen/
    invoice_cn/proofread do not carry not_implemented_routes. This guards
    against accidentally introducing stale honesty declarations there.
    """
    for pkg in (
        "docwen_plugin_document",
        "docwen_plugin_presentation",
        "docwen_plugin_spreadsheet",
        "docwen_plugin_markup",
        "docwen_plugin_image",
        "docwen_plugin_optimizer_gongwen",
        "docwen_plugin_optimizer_invoice_cn",
        "docwen_plugin_proofread",
    ):
        extra = _load_manifest(pkg).extra
        assert "not_implemented_routes" not in extra, (
            f"{pkg} unexpectedly declares not_implemented_routes; "
            f"if real, add it to _HONESTY_PLUGINS and a targeted test."
        )


def test_markdown_route_schemas_declare_app_resolved_metadata_options() -> None:
    """Markdown-producing route schemas should declare metadata/i18n request keys they consume."""
    from docwen_plugin_document.manifest import DOCX_TO_MD_OPTIONS_SCHEMA
    from docwen_plugin_image.manifest import IMAGE_TO_MD_OPTIONS_SCHEMA
    from docwen_plugin_layout.manifest import LAYOUT_TO_MD_OPTIONS_SCHEMA
    from docwen_plugin_markup.manifest import _MARKUP_OPTIONS_SCHEMA
    from docwen_plugin_optimizer_gongwen.manifest import GONGWEN_OPTIONS_SCHEMA
    from docwen_plugin_optimizer_invoice_cn.manifest import INVOICE_CN_OPTIONS_SCHEMA
    from docwen_plugin_presentation.manifest import PPTX_TO_MD_OPTIONS_SCHEMA
    from docwen_plugin_spreadsheet.manifest import SPREADSHEET_TO_MD_OPTIONS_SCHEMA

    metadata_schemas = (
        DOCX_TO_MD_OPTIONS_SCHEMA,
        IMAGE_TO_MD_OPTIONS_SCHEMA,
        LAYOUT_TO_MD_OPTIONS_SCHEMA,
        _MARKUP_OPTIONS_SCHEMA,
        INVOICE_CN_OPTIONS_SCHEMA,
        PPTX_TO_MD_OPTIONS_SCHEMA,
        SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
    )

    for schema in metadata_schemas:
        properties = schema["properties"]
        assert properties["locale"]["type"] == "string"
        assert properties["locale"]["x-docwen-status"] == "implemented"
        assert properties["yaml_key_labels"]["type"] == "object"
        assert properties["yaml_key_labels"]["additionalProperties"] == {"type": "string"}
        assert properties["yaml_key_labels"]["x-docwen-status"] == "implemented"

    gongwen_properties = GONGWEN_OPTIONS_SCHEMA["properties"]
    assert gongwen_properties["locale"]["type"] == "string"
    assert gongwen_properties["locale"]["x-docwen-status"] == "implemented"
    assert "yaml_key_labels" not in gongwen_properties


def test_option_manifest_static_scan_candidates_keep_focused_schema_guards() -> None:
    """Near-miss option/manifest scan candidates should stay covered by focused tests."""
    image_tests = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "tests" / "test_plugin_image_imports.py"
    ).read_text(encoding="utf-8")
    presentation_tests = (
        PROJECT_ROOT / "packages" / "plugins" / "presentation" / "tests" / "test_plugin_presentation_imports.py"
    ).read_text(encoding="utf-8")
    proofread_tests = (
        PROJECT_ROOT / "packages" / "plugins" / "proofread" / "tests" / "test_proofread_non_docx.py"
    ).read_text(encoding="utf-8")
    gongwen_tests = (
        PROJECT_ROOT / "packages" / "plugins" / "optimizers" / "gongwen" / "tests" / "test_plugin_gongwen_imports.py"
    ).read_text(encoding="utf-8")
    document_tests = (
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_plugin_document_imports.py"
    ).read_text(encoding="utf-8")

    assert "test_image_manifest_keeps_route_specific_options_scoped" in image_tests
    for token in ("quality_mode", "target_format", "compress_mode", "size_unit"):
        assert token in image_tests

    assert "test_presentation_to_md_manifest_declares_consumed_markdown_export_options" in presentation_tests
    for token in ("export_notes", "yaml_key_labels", "image_mode", "ocr_placement", "image_link_style"):
        assert token in presentation_tests

    assert "test_manifest_declares_only_consumed_proofread_action_options" in proofread_tests
    for token in ("skip_code_blocks", "skip_quote_blocks", "enable_symbol_pairing", "proofread_rules"):
        assert token in proofread_tests

    assert "test_gongwen_manifest_declares_only_consumed_action_options" in gongwen_tests
    for token in ("to_md_enable_ocr", "numbering_scheme", "output_dir", "yaml_key_labels"):
        assert token in gongwen_tests
    assert "Legacy alias for to_md_enable_ocr" not in gongwen_tests

    assert "test_docx_to_md_manifest_declares_consumed_markdown_export_options" in document_tests
    for token in ("code_block_style_aliases", "quote_style_aliases", "quote_generic_names"):
        assert token in document_tests


def test_option_manifest_static_scan_candidate_source_keys_match_schemas() -> None:
    """Source-level request-key reads should stay aligned with route schemas."""
    from docwen_plugin_document.manifest import DOCX_TO_MD_OPTIONS_SCHEMA
    from docwen_plugin_image.manifest import (
        EXPLICIT_IMAGE_FORMAT_OPTIONS_SCHEMA,
        IMAGE_FORMAT_OPTIONS_SCHEMA,
        IMAGE_MERGE_OPTIONS_SCHEMA,
        IMAGE_TO_PDF_OPTIONS_SCHEMA,
    )
    from docwen_plugin_optimizer_gongwen.manifest import GONGWEN_OPTIONS_SCHEMA
    from docwen_plugin_presentation.manifest import PPTX_TO_MD_OPTIONS_SCHEMA
    from docwen_plugin_proofread.rules import PROOFREAD_OPTIONS_SCHEMA

    image_pdf_source = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "src" / "docwen_plugin_image" / "to_pdf" / "converter.py"
    ).read_text(encoding="utf-8")
    image_format_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "image"
        / "src"
        / "docwen_plugin_image"
        / "format_conversion"
        / "converter.py"
    ).read_text(encoding="utf-8")
    image_merge_source = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "src" / "docwen_plugin_image" / "merge" / "converter.py"
    ).read_text(encoding="utf-8")
    image_common_source = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "src" / "docwen_plugin_image" / "_common.py"
    ).read_text(encoding="utf-8")

    assert 'options.get("pdf_quality")' not in image_pdf_source
    assert 'options.get("quality_mode")' in image_pdf_source
    assert 'options.get("size_limit")' not in image_pdf_source
    assert 'options.get("compress_mode")' not in image_pdf_source
    assert set(IMAGE_TO_PDF_OPTIONS_SCHEMA["properties"]) == {"quality_mode"}
    assert 'context.request.options.get("target_format")' in image_format_source
    assert "target_format" in IMAGE_FORMAT_OPTIONS_SCHEMA["properties"]
    assert "target_format" not in EXPLICIT_IMAGE_FORMAT_OPTIONS_SCHEMA["properties"]
    assert 'options.get("compress_mode")' in image_common_source
    assert 'options.get("size_limit")' in image_common_source
    assert 'options.get("size_unit", "KB")' in image_common_source
    assert {"compress_mode", "size_limit", "size_unit"} <= set(IMAGE_FORMAT_OPTIONS_SCHEMA["properties"])
    assert 'options.get("mode")' in image_merge_source
    assert 'options.get("keep_alpha")' in image_merge_source
    assert {"mode", "keep_alpha"} <= set(IMAGE_MERGE_OPTIONS_SCHEMA["properties"])

    presentation_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "presentation"
        / "src"
        / "docwen_plugin_presentation"
        / "pptx_md"
        / "converter.py"
    ).read_text(encoding="utf-8")
    presentation_policy_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "presentation"
        / "src"
        / "docwen_plugin_presentation"
        / "pptx_md"
        / "request_policy.py"
    ).read_text(encoding="utf-8")
    presentation_option_sources = presentation_source + "\n" + presentation_policy_source
    presentation_properties = PPTX_TO_MD_OPTIONS_SCHEMA["properties"]
    for key in (
        "export_notes",
        "to_md_keep_images",
        "to_md_enable_ocr",
        "ocr_language",
        "locale",
        "image_mode",
        "ocr_placement",
        "image_link_style",
        "yaml_key_labels",
    ):
        assert any(
            token in presentation_option_sources
            for token in (
                f'opts.get("{key}")',
                f'opts.get("{key}",',
                f'options.get("{key}")',
                f'options.get("{key}",',
            )
        )
        assert key in presentation_properties

    proofread_common_source = (
        PROJECT_ROOT / "packages" / "plugins" / "proofread" / "src" / "docwen_plugin_proofread" / "_common.py"
    ).read_text(encoding="utf-8")
    proofread_skip_source = (
        PROJECT_ROOT / "packages" / "plugins" / "proofread" / "src" / "docwen_plugin_proofread" / "skip_policy.py"
    ).read_text(encoding="utf-8")
    proofread_properties = PROOFREAD_OPTIONS_SCHEMA["properties"]
    for key in (
        "enable_symbol_pairing",
        "enable_symbol_correction",
        "enable_typos_rule",
        "enable_sensitive_word",
    ):
        assert f'engine.get("{key}"' in proofread_common_source
        assert key in proofread_properties
    for key in ("skip_code_blocks", "skip_quote_blocks"):
        assert key in proofread_skip_source
        assert key in proofread_properties
    assert "symbol_pairing" not in proofread_properties
    assert "proofread_rules" not in proofread_properties

    gongwen_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "optimizers"
        / "gongwen"
        / "src"
        / "docwen_plugin_optimizer_gongwen"
        / "pipeline.py"
    ).read_text(encoding="utf-8")
    gongwen_properties = GONGWEN_OPTIONS_SCHEMA["properties"]
    for key in (
        "remove_numbering",
        "add_numbering",
        "numbering_scheme",
        "to_md_enable_ocr",
        "ocr_language",
        "locale",
    ):
        assert key in gongwen_source
        assert key in gongwen_properties
    assert 'options.get("output_dir")' in gongwen_source
    assert "output_dir" not in gongwen_properties
    assert "yaml_key_labels" not in gongwen_properties

    document_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "to_markdown"
        / "converter.py"
    ).read_text(encoding="utf-8")
    document_policy_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "to_markdown"
        / "request_policy.py"
    ).read_text(encoding="utf-8")
    document_properties = DOCX_TO_MD_OPTIONS_SCHEMA["properties"]
    for key in ("code_block_style_aliases", "quote_style_aliases", "quote_generic_names"):
        assert (
            f'options.get("{key}"' in document_source or f'_option_strings(options, "{key}")' in document_policy_source
        )
        assert key in document_properties


def test_remaining_plugin_source_option_reads_match_route_schemas() -> None:
    """Remaining plugin-family request-key reads should stay route-schema scoped."""
    from docwen_plugin_layout.manifest import (
        LAYOUT_TO_IMAGE_OPTIONS_SCHEMA,
        LAYOUT_TO_MD_OPTIONS_SCHEMA,
        PDF_OPERATION_OPTIONS_SCHEMA,
        PDF_SPLIT_OPTIONS_SCHEMA,
    )
    from docwen_plugin_markdown.manifest import (
        MD_NUMBERING_OPTIONS_SCHEMA,
        MD_TO_DOCX_OPTIONS_SCHEMA,
        MD_TO_SPREADSHEET_TEMPLATE_OPTIONS_SCHEMA,
    )
    from docwen_plugin_markup.manifest import _MARKUP_OPTIONS_SCHEMA
    from docwen_plugin_spreadsheet.manifest import (
        SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
        TABLE_MERGE_OPTIONS_SCHEMA,
    )

    layout_md_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "layout"
        / "src"
        / "docwen_plugin_layout"
        / "to_markdown"
        / "converter.py"
    ).read_text(encoding="utf-8")
    layout_image_source = (
        PROJECT_ROOT / "packages" / "plugins" / "layout" / "src" / "docwen_plugin_layout" / "to_image" / "converter.py"
    ).read_text(encoding="utf-8")
    layout_operations_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "layout"
        / "src"
        / "docwen_plugin_layout"
        / "operations"
        / "converter.py"
    ).read_text(encoding="utf-8")

    layout_md_properties = LAYOUT_TO_MD_OPTIONS_SCHEMA["properties"]
    for key in (
        "to_md_keep_images",
        "to_md_enable_ocr",
        "render_dpi",
        "image_mode",
        "ocr_language",
        "locale",
        "image_link_style",
        "yaml_key_labels",
    ):
        assert f'options.get("{key}"' in layout_md_source
        assert key in layout_md_properties
    assert 'options.get("render_dpi"' in layout_image_source
    assert set(LAYOUT_TO_IMAGE_OPTIONS_SCHEMA["properties"]) == {"render_dpi"}
    assert {"split_mode", "pages"} <= set(PDF_SPLIT_OPTIONS_SCHEMA["properties"])
    assert 'options.get("split_mode"' in layout_operations_source
    assert 'options.get("pages"' in layout_operations_source
    assert PDF_OPERATION_OPTIONS_SCHEMA["properties"] == {}

    markup_sources = [
        (
            PROJECT_ROOT / "packages" / "plugins" / "markup" / "src" / "docwen_plugin_markup" / family / "converter.py"
        ).read_text(encoding="utf-8")
        for family in ("web_archive", "publication", "note_export")
    ]
    markup_properties = _MARKUP_OPTIONS_SCHEMA["properties"]
    for key in (
        "to_md_keep_images",
        "to_md_enable_ocr",
        "image_mode",
        "ocr_placement",
        "ocr_language",
        "locale",
        "image_link_style",
        "yaml_key_labels",
    ):
        assert all(f'options.get("{key}"' in source for source in markup_sources)
        assert key in markup_properties

    spreadsheet_md_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "spreadsheet"
        / "src"
        / "docwen_plugin_spreadsheet"
        / "to_markdown"
        / "converter.py"
    ).read_text(encoding="utf-8")
    spreadsheet_merge_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "spreadsheet"
        / "src"
        / "docwen_plugin_spreadsheet"
        / "table_merger"
        / "converter.py"
    ).read_text(encoding="utf-8")

    spreadsheet_md_properties = SPREADSHEET_TO_MD_OPTIONS_SCHEMA["properties"]
    for key in (
        "to_md_keep_images",
        "table_merge_strategy",
        "yaml_key_labels",
        "to_md_enable_ocr",
        "image_mode",
        "ocr_placement",
        "ocr_language",
        "locale",
        "image_link_style",
    ):
        assert f'get("{key}"' in spreadsheet_md_source
        assert key in spreadsheet_md_properties
    spreadsheet_merge_properties = TABLE_MERGE_OPTIONS_SCHEMA["properties"]
    assert 'options.get("merge_mode"' in spreadsheet_merge_source
    assert 'options.get("offset_range"' in spreadsheet_merge_source
    assert set(spreadsheet_merge_properties) == {"merge_mode", "offset_range"}
    assert "base_table" not in spreadsheet_merge_properties

    markdown_numbering_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "markdown"
        / "src"
        / "docwen_plugin_markdown"
        / "numbering"
        / "converter.py"
    ).read_text(encoding="utf-8")
    markdown_docx_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "markdown"
        / "src"
        / "docwen_plugin_markdown"
        / "to_docx"
        / "converter.py"
    ).read_text(encoding="utf-8")
    markdown_spreadsheet_source = (
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "markdown"
        / "src"
        / "docwen_plugin_markdown"
        / "to_spreadsheet"
        / "converter.py"
    ).read_text(encoding="utf-8")

    numbering_properties = MD_NUMBERING_OPTIONS_SCHEMA["properties"]
    for key in ("remove_numbering", "add_numbering", "numbering_scheme"):
        assert f'options.get("{key}"' in markdown_numbering_source
        assert key in numbering_properties
    assert "heading_numbering_render_mode" not in numbering_properties

    docx_properties = MD_TO_DOCX_OPTIONS_SCHEMA["properties"]
    assert docx_properties["locale"]["enum"] == list(SHIPPED_STYLE_LOCALES)
    for key in (
        "formatting_mode",
        "remove_numbering",
        "add_numbering",
        "heading_numbering_render_mode",
        "numbering_scheme",
        "template_name",
        "hr_mapping",
    ):
        assert f'options.get("{key}"' in markdown_docx_source
        assert key in docx_properties
    assert "list_separator" not in docx_properties
    assert '"conversion.md_to_docx.list_separator"' in markdown_docx_source
    assert '"conversion.md_to_docx.list_separator"' in markdown_spreadsheet_source

    spreadsheet_template_properties = MD_TO_SPREADSHEET_TEMPLATE_OPTIONS_SCHEMA["properties"]
    assert 'options", {}).get("template_name"' in markdown_spreadsheet_source
    assert set(spreadsheet_template_properties) == {"template_name"}
