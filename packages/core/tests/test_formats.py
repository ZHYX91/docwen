"""Contract tests for format categories and route registry."""

from __future__ import annotations

import pytest

from docwen_core.formats.categories import (
    CATEGORY_DOCUMENT,
    CATEGORY_IMAGE,
    CATEGORY_LAYOUT,
    CATEGORY_MARKDOWN,
    CATEGORY_MARKUP,
    CATEGORY_PRESENTATION,
    CATEGORY_SPREADSHEET,
    FORMAT_CATEGORY,
    get_category,
    get_media_type,
)
from docwen_core.formats.routes import RouteRegistry
from docwen_core.models.manifest import RouteSpec

pytestmark = pytest.mark.contract


class TestFormatCategories:
    def test_known_formats(self) -> None:
        assert get_category("docx") == CATEGORY_DOCUMENT
        assert get_category("xlsx") == CATEGORY_SPREADSHEET
        assert get_category("png") == CATEGORY_IMAGE
        assert get_category("pdf") == CATEGORY_LAYOUT
        assert get_category("md") == CATEGORY_MARKDOWN

    def test_case_insensitive(self) -> None:
        assert get_category("DOCX") == CATEGORY_DOCUMENT
        assert get_category("Pdf") == CATEGORY_LAYOUT

    def test_markup_formats(self) -> None:
        assert get_category("html") == CATEGORY_MARKUP
        assert get_category("epub") == CATEGORY_MARKUP
        assert get_category("enex") == CATEGORY_MARKUP

    def test_presentation_formats(self) -> None:
        assert get_category("pptx") == CATEGORY_PRESENTATION
        assert get_category("ppt") == CATEGORY_PRESENTATION

    def test_unknown_format_returns_other(self) -> None:
        assert get_category("xyz_unknown") == "other"

    @pytest.mark.parametrize("retired_format", ["caj", "oxps"])
    def test_retired_unsupported_formats_are_not_recognized(self, retired_format: str) -> None:
        assert retired_format not in FORMAT_CATEGORY
        assert get_category(retired_format) == "other"
        assert get_media_type(retired_format) == "application/octet-stream"

    def test_format_category_mapping_covers_route_matrix_formats(self) -> None:
        """Every format in the route_matrix.csv should have a defined category."""
        # Source formats from route_matrix.csv
        route_formats = [
            "document",  # abstract category, not a format
            "spreadsheet",
            "image",
            "layout",
            "markdown",
            "html",
            "mhtml",
            "htm",
            "mht",
            "enex",
            "pptx",
            "ppt",
            "epub",
            "docx",
            "doc",
            "odt",
            "rtf",
            "wps",
            "xlsx",
            "xls",
            "ods",
            "et",
            "csv",
            "tsv",
            "ofd",
            "xps",
        ]
        missing = [
            f
            for f in route_formats
            if f not in FORMAT_CATEGORY
            and f
            not in {
                "document",
                "spreadsheet",
                "image",
                "layout",
                "markdown",
            }
        ]
        # Abstract categories are not real formats
        assert missing == []

    def test_media_types_for_common_formats(self) -> None:
        assert get_media_type("docx") == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        assert get_media_type("md") == "text/markdown"
        assert get_media_type("pdf") == "application/pdf"
        assert get_media_type("png") == "image/png"
        assert get_media_type("xlsx") == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class TestRouteRegistry:
    def test_register_and_find_exact(self) -> None:
        reg = RouteRegistry()
        reg.register(RouteSpec(source_format="markdown", target_format="docx"), "plugin-md")
        entry = reg.find("markdown", "docx")
        assert entry is not None
        assert entry.plugin_id == "plugin-md"

    def test_find_none_for_unknown_route(self) -> None:
        reg = RouteRegistry()
        assert reg.find("unknown", "xyz") is None

    def test_action_name_precedence(self) -> None:
        """Exact action match should beat empty action match."""
        reg = RouteRegistry()
        reg.register(RouteSpec(source_format="docx", target_format="md", action_name=""), "plugin-generic")
        reg.register(RouteSpec(source_format="docx", target_format="md", action_name="validate"), "plugin-validate")
        entry = reg.find("docx", "md", action_name="validate")
        assert entry is not None
        assert entry.plugin_id == "plugin-validate"

    def test_list_routes(self) -> None:
        reg = RouteRegistry()
        reg.register(RouteSpec(source_format="a", target_format="b"), "p1")
        reg.register(RouteSpec(source_format="c", target_format="d"), "p2")
        assert len(reg.list_routes()) == 2

    def test_list_plugin_routes(self) -> None:
        reg = RouteRegistry()
        reg.register(RouteSpec(source_format="a", target_format="b"), "p1")
        reg.register(RouteSpec(source_format="c", target_format="d"), "p1")
        reg.register(RouteSpec(source_format="e", target_format="f"), "p2")
        p1_routes = reg.list_plugin_routes("p1")
        assert len(p1_routes) == 2
        assert len(reg.list_plugin_routes("nonexistent")) == 0

    def test_route_registry_can_express_md_to_docx(self) -> None:
        """Route registry must be able to register and resolve ROUTE-MD-DOCX-001."""
        reg = RouteRegistry()
        reg.register(
            RouteSpec(
                source_format="markdown",
                target_format="docx",
                action_name="",
                label="Markdown → DOCX",
            ),
            "docwen_plugin_markdown",
        )
        entry = reg.find("markdown", "docx")
        assert entry is not None
        assert entry.plugin_id == "docwen_plugin_markdown"
        assert entry.route.source_format == "markdown"
        assert entry.route.target_format == "docx"
