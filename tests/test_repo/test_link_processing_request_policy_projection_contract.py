"""Fail-closed evidence guards for VIS-2026-07-19-140 Link request policy."""

from __future__ import annotations

from dataclasses import fields
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "link-processing-request-policy-2026-07-19.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def test_link_runtime_is_a_twelve_field_request_snapshot_without_global_fallback() -> None:
    from docwen_core.export_semantics import LinkRuntimeConfig

    assert tuple(field.name for field in fields(LinkRuntimeConfig)) == (
        "max_depth",
        "non_embed_wiki_mode",
        "non_embed_markdown_mode",
        "embed_wiki_image_mode",
        "embed_markdown_image_mode",
        "embed_md_file_mode",
        "search_dirs",
        "detect_circular",
        "file_not_found_mode",
        "circular_reference_mode",
        "max_depth_reached_mode",
        "auto_link_bare_url",
    )
    assert LinkRuntimeConfig().auto_link_bare_url is False

    production_sources = "\n".join(
        path.read_text(encoding="utf-8")
        for path in (PROJECT_ROOT / "packages").rglob("*.py")
        if "/src/" in path.relative_to(PROJECT_ROOT).as_posix()
    )
    assert "configure_link_runtime_config" not in production_sources
    assert "_get_link_cfg" not in production_sources
    assert production_sources.count("process_markdown_links(") >= 3
    assert production_sources.count("link_config=") >= 3


def test_markdown_routes_keep_scoped_marker_and_bare_url_ownership_explicit() -> None:
    docx = _read("packages/plugins/markdown/src/docwen_plugin_markdown/to_docx/converter.py")
    spreadsheet = _read("packages/plugins/markdown/src/docwen_plugin_markdown/to_spreadsheet/converter.py")
    parser = _read("packages/plugins/markdown/src/docwen_plugin_markdown/mistune_extensions.py")
    core_regression = _read("packages/core/tests/test_link_runtime_processing_contract_*.py")
    route_regression = _read("packages/plugins/markdown/tests/test_link_processing_routes_*.py")

    assert "image_scope = secrets.token_urlsafe(24)" in docx
    assert 'target_format="docx"' in docx
    assert "materialize_image_placeholders(" in docx and "image_scope=image_scope" in docx
    assert "parse_markdown_text(md_body, auto_link_bare_url=False)" in docx

    assert spreadsheet.count("image_scope = secrets.token_urlsafe(24)") >= 2
    assert 'target_format="xlsx"' in spreadsheet
    assert 'target_format="csv"' in spreadsheet
    assert "markdown_body = _csv_image_fallbacks(" in spreadsheet
    assert spreadsheet.count("image_scope=image_scope") >= 6

    assert "if auto_link_bare_url:" in parser
    assert "test_docx_remove_policy_preserves_literal_unscoped_image_marker" in route_regression
    assert "test_xlsx_literal_and_inline_image_markers_do_not_create_drawings" in route_regression
    assert "test_csv_literal_and_inline_image_markers_stay_literal" in route_regression
    assert "test_link_rewrites_protect_strict_and_indented_code_blocks" in core_regression

    markdown_sources = "\n".join(
        path.read_text(encoding="utf-8")
        for path in (PROJECT_ROOT / "packages" / "plugins" / "markdown" / "src").rglob("*.py")
    )
    assert "docwen_core.links._" not in markdown_sources
    core_links_api = _read("packages/core/src/docwen_core/links/__init__.py")
    for public_adapter in (
        "make_table_safe",
        "split_markdown_block_segments",
        "split_markdown_inline_segments",
    ):
        assert f'"{public_adapter}"' in core_links_api
