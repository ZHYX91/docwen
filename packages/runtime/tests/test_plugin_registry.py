"""Tests for PluginRegistry."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit
from tests.support.plugin import FakePlugin

from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_runtime.plugin_registry.registry import PluginRegistry


@pytest.fixture
def md_plugin() -> FakePlugin:
    return FakePlugin(
        PluginManifest(
            plugin_id="docwen_plugin_markdown",
            name="Markdown Plugin",
            version="0.1.0",
            routes=[
                RouteSpec(source_format="markdown", target_format="docx", label="Markdown → DOCX"),
                RouteSpec(source_format="markdown", target_format="xlsx", label="Markdown → XLSX"),
            ],
        )
    )


@pytest.fixture
def docx_plugin() -> FakePlugin:
    return FakePlugin(
        PluginManifest(
            plugin_id="docwen_plugin_document",
            name="Document Plugin",
            version="0.1.0",
            routes=[
                RouteSpec(source_format="docx", target_format="md", label="DOCX → MD"),
            ],
        )
    )


class TestPluginRegistry:
    def test_register_and_get(self, md_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)
        assert reg.get("docwen_plugin_markdown") is md_plugin
        assert len(reg) == 1

    def test_register_multiple(self, md_plugin: FakePlugin, docx_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)
        reg.register(docx_plugin)
        assert len(reg) == 2
        assert "docwen_plugin_markdown" in reg
        assert "docwen_plugin_document" in reg

    def test_register_replaces_duplicate(self, md_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)
        # Register another plugin with the same id
        md2 = FakePlugin(
            PluginManifest(
                plugin_id="docwen_plugin_markdown",
                name="Markdown Plugin v2",
                version="0.2.0",
                routes=[RouteSpec(source_format="markdown", target_format="pdf")],
            )
        )
        reg.register(md2)
        assert len(reg) == 1
        assert reg.get("docwen_plugin_markdown") is md2

    def test_unregister(self, md_plugin: FakePlugin, docx_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)
        reg.register(docx_plugin)
        reg.unregister("docwen_plugin_markdown")
        assert len(reg) == 1
        assert "docwen_plugin_markdown" not in reg
        assert "docwen_plugin_document" in reg

    def test_unregister_nonexistent_noop(self) -> None:
        reg = PluginRegistry()
        reg.unregister("nonexistent")  # no error

    def test_find_plugin_by_route(self, md_plugin: FakePlugin, docx_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)
        reg.register(docx_plugin)

        found = reg.find_plugin("markdown", "docx")
        assert found is md_plugin

        found = reg.find_plugin("docx", "md")
        assert found is docx_plugin

        found = reg.find_plugin("markdown", "pdf")
        assert found is None

    def test_find_manifest_by_route(self, md_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)

        m = reg.find_manifest("markdown", "docx")
        assert m is not None
        assert m.plugin_id == "docwen_plugin_markdown"

    def test_plugin_ids(self, md_plugin: FakePlugin, docx_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)
        reg.register(docx_plugin)
        ids = reg.plugin_ids
        assert sorted(ids) == ["docwen_plugin_document", "docwen_plugin_markdown"]

    def test_list_manifests(self, md_plugin: FakePlugin, docx_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)
        manifests = reg.list_manifests()
        assert len(manifests) == 1
        assert manifests[0].plugin_id == "docwen_plugin_markdown"

    def test_route_registry_updated_on_unregister(self, md_plugin: FakePlugin, docx_plugin: FakePlugin) -> None:
        reg = PluginRegistry()
        reg.register(md_plugin)
        reg.register(docx_plugin)

        # Before unregister, both routes exist
        assert reg.find_plugin("markdown", "docx") is not None
        assert reg.find_plugin("docx", "md") is not None

        reg.unregister("docwen_plugin_markdown")

        # After unregister, only docx plugin remains
        assert reg.find_plugin("markdown", "docx") is None
        assert reg.find_plugin("docx", "md") is docx_plugin
