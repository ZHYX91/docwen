"""Tests for RouteResolver."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit
from tests.support.plugin import FakePlugin

from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_runtime.engine.route_resolver import RouteResolutionError, RouteResolver
from docwen_runtime.plugin_registry.registry import PluginRegistry


@pytest.fixture
def resolver() -> RouteResolver:
    reg = PluginRegistry()
    reg.register(
        FakePlugin(
            PluginManifest(
                plugin_id="p1",
                name="P1",
                version="1",
                routes=[
                    RouteSpec(source_format="markdown", target_format="docx", label="MD→DOCX"),
                    RouteSpec(source_format="markdown", target_format="md", label="MD→MD"),
                    RouteSpec(source_format="markdown", target_format="xlsx", label="MD→XLSX"),
                ],
            )
        )
    )
    reg.register(
        FakePlugin(
            PluginManifest(
                plugin_id="p2",
                name="P2",
                version="1",
                routes=[
                    RouteSpec(source_format="docx", target_format="md", label="DOCX→MD"),
                    RouteSpec(source_format="pdf", target_format="md", action_name="", label="PDF→MD"),
                ],
            )
        )
    )
    return RouteResolver(reg)


class TestRouteResolver:
    def test_resolve_success(self, resolver: RouteResolver) -> None:
        plugin_id, route = resolver.resolve(
            FileRef(path="/f.md", format="markdown", category="markdown"),
            "docx",
        )
        assert plugin_id == "p1"
        assert route.source_format == "markdown"
        assert route.target_format == "docx"

    def test_resolve_different_plugin(self, resolver: RouteResolver) -> None:
        plugin_id, _route = resolver.resolve(
            FileRef(path="/f.docx", format="docx", category="document"),
            "md",
        )
        assert plugin_id == "p2"

    def test_resolve_no_match_raises(self, resolver: RouteResolver) -> None:
        with pytest.raises(RouteResolutionError) as exc_info:
            resolver.resolve(
                FileRef(path="/f.md", format="markdown", category="markdown"),
                "pdf",
            )
        assert "markdown" in str(exc_info.value)
        assert "pdf" in str(exc_info.value)

    def test_resolve_plugin_id_convenience(self, resolver: RouteResolver) -> None:
        pid = resolver.resolve_plugin_id(
            FileRef(path="/f.md", format="markdown", category="markdown"),
            "docx",
        )
        assert pid == "p1"

    def test_concrete_format_falls_back_to_frozen_workflow_category(self, resolver: RouteResolver) -> None:
        plugin_id, route = resolver.resolve(
            FileRef(path="/looks-like-document.docx", format="txt", category="markdown"),
            "docx",
        )

        assert plugin_id == "p1"
        assert route.source_format == "markdown"

    def test_concrete_format_precedes_conflicting_category_and_suffix(self, resolver: RouteResolver) -> None:
        plugin_id, route = resolver.resolve(
            FileRef(path="/looks-like-markdown.md", format="docx", category="markdown"),
            "md",
        )

        assert plugin_id == "p2"
        assert route.source_format == "docx"

    def test_unknown_format_may_use_frozen_category_without_reinspection(self, resolver: RouteResolver) -> None:
        plugin_id, route = resolver.resolve(
            FileRef(path="/looks-like-layout.pdf", format="unknown", category="markdown"),
            "docx",
        )

        assert plugin_id == "p1"
        assert route.source_format == "markdown"

    def test_unknown_format_does_not_reinspect_or_guess_from_suffix(
        self,
        resolver: RouteResolver,
        tmp_path: Path,
    ) -> None:
        recognizable = tmp_path / "looks-like-markdown.md"
        recognizable.write_text("# Content that a legacy resolver would detect", encoding="utf-8")

        with pytest.raises(RouteResolutionError) as exc_info:
            resolver.resolve(
                FileRef(path=str(recognizable), format="unknown", category="other"),
                "docx",
            )

        assert exc_info.value.source_format == "unknown"

    def test_resolution_error_contains_formats(self) -> None:
        with pytest.raises(RouteResolutionError) as exc_info:
            raise RouteResolutionError("jpg", "docx", "")
        err = exc_info.value
        assert err.source_format == "jpg"
        assert err.target_format == "docx"
        assert err.action_name == ""
