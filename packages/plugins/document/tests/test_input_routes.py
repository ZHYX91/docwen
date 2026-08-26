"""Document family input route verification tests.

Verifies that DocumentPlugin only handles document-family routes
(docx/document → md, SmartConverter interconversion) and correctly
rejects routes that have been migrated to markup/presentation/print.
"""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


# ═══════════════════════════════════════════════════════════════════════════
# can_handle() — document family routes
# ═══════════════════════════════════════════════════════════════════════════


class TestDocumentFamilyRoutes:
    """Verify document-family routes are accepted by DocumentPlugin."""

    def test_docx_to_md_accepted(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("docx", "md") is True

    def test_document_to_md_accepted(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("document", "md") is True

    def test_markdown_target_identifier_is_rejected(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("docx", "markdown") is False

    @pytest.mark.parametrize("target", ["doc", "odt", "rtf", "wps"])
    def test_docx_to_document_formats_accepted(self, target: str) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("docx", target) is True

    @pytest.mark.parametrize("target", ["docx", "odt", "rtf", "wps"])
    def test_doc_to_document_formats_accepted(self, target: str) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("doc", target) is True

    def test_gongwen_submode_rejected(self) -> None:
        """Gongwen is now handled by the dedicated optimizer plugin."""
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("docx", "md", "gongwen") is False

    def test_invoice_cn_submode_rejected(self) -> None:
        """Invoice CN is now handled by the dedicated optimizer plugin."""
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("docx", "md", "invoice_cn") is False


# ═══════════════════════════════════════════════════════════════════════════
# can_handle() — migrated routes MUST be rejected
# ═══════════════════════════════════════════════════════════════════════════


class TestMigratedRoutesRejected:
    """Routes migrated to markup/presentation/print must be rejected."""

    def test_enex_to_md_rejected(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("enex", "md") is False

    def test_html_to_md_rejected(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("html", "md") is False

    def test_mhtml_to_md_rejected(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("mhtml", "md") is False

    def test_pptx_to_md_rejected(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("pptx", "md") is False

    def test_ppt_to_md_rejected(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("ppt", "md") is False

    def test_document_to_pdf_rejected(self) -> None:
        """Document→PDF belongs to print plugin, not document."""
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert plugin.can_handle("document", "pdf") is False


# ═══════════════════════════════════════════════════════════════════════════
# Manifest structure
# ═══════════════════════════════════════════════════════════════════════════


class TestManifestStructure:
    """Verify manifest structure matches document-only responsibilities."""

    def test_manifest_has_22_routes(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert len(plugin.manifest.routes) == 22, (
            f"Expected 22 routes (docx→md, document→md, 20 SmartConverter), got {len(plugin.manifest.routes)}"
        )

    def test_no_markup_routes_in_manifest(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        manifest_sources = {r.source_format for r in plugin.manifest.routes}
        markup_formats = {"html", "htm", "mhtml", "mht", "enex", "epub"}
        assert manifest_sources.isdisjoint(markup_formats), (
            f"Document manifest should not contain markup formats: {manifest_sources & markup_formats}"
        )

    def test_no_presentation_routes_in_manifest(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        manifest_sources = {r.source_format for r in plugin.manifest.routes}
        presentation_formats = {"pptx", "ppt"}
        assert manifest_sources.isdisjoint(presentation_formats), (
            f"Document manifest should not contain presentation formats: {manifest_sources & presentation_formats}"
        )
