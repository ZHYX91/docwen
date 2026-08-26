"""Concrete parser and Application-preconversion proofread contracts.

The plugin itself parses DOCX and Markdown only.  Runtime discovery exposes a
category-level document action so the Application can pre-convert legacy
DOC/WPS/RTF/ODT content before the plugin receives it.

These tests verify that:
1. The manifest no longer declares routes for non-DOCX formats.
2. DOCX and Markdown validation still work correctly.
3. Helper functions in _common.py that remain are still functional.
"""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.golden


# ── Helpers ──────────────────────────────────────────────────────────────


def _create_minimal_docx(path: str, paragraphs: list[str]) -> None:
    """Create a minimal DOCX file with the given paragraph texts."""
    from docx import Document

    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)


def _build_fake_context(
    input_path: str,
    staging_dir: str,
    *,
    target_format: str = "docx",
    action_name: str = "validate",
    source_format: str = "docx",
    options: dict[str, Any] | None = None,
    pre_cancelled: bool = False,
    config_snapshot: dict[str, Any] | None = None,
) -> Any:
    """Build a minimal fake PluginExecutionContext for proofread testing."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    file_refs = [
        FileRef(
            path=input_path,
            format=source_format,
            category="document",
        )
    ]
    request = ConversionRequest(
        request_id="test-proofread-nondocx-001",
        input_refs=file_refs,
        target_format=target_format,
        action_name=action_name,
        options=options or {},
        output_policy=OutputPolicy(),
        config_snapshot=config_snapshot or {},
    )
    config = FakeConfigView(values=config_snapshot or {})
    token = CancellationToken()
    if pre_cancelled:
        token.cancel("test cancellation")
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle(input_path, staging_dir),
        config,
        FakeProgressSink(),
        token,
        FakePluginLogger(),
    )


# ═══════════════════════════════════════════════════════════════════════════
# Manifest tests — non-DOCX routes are no longer declared
# ═══════════════════════════════════════════════════════════════════════════


@pytest.mark.contract
class TestNonDocxRoutes:
    """Verify direct parsers stay strict while the composed route is declared."""

    def test_can_handle_rejects_non_docx(self) -> None:
        """The plugin must not parse non-DOCX formats directly."""
        from docwen_plugin_proofread import ProofreadPlugin

        plugin = ProofreadPlugin()
        for fmt in ("doc", "odt", "rtf", "wps"):
            assert plugin.can_handle(fmt, "docx", "validate") is False, (
                f"Expected can_handle({fmt!r}) to be False — only concrete DOCX is admitted"
            )

    def test_can_handle_docx_still_works(self) -> None:
        """The original DOCX route must still be matched."""
        from docwen_plugin_proofread import ProofreadPlugin

        plugin = ProofreadPlugin()
        assert plugin.can_handle("docx", "docx", "validate") is True

    def test_can_handle_markdown_still_works(self) -> None:
        """The Markdown route must still be matched."""
        from docwen_plugin_proofread import ProofreadPlugin

        plugin = ProofreadPlugin()
        assert plugin.can_handle("markdown", "markdown", "validate") is True

    def test_can_handle_rejects_unsupported(self) -> None:
        """Unsupported formats must be rejected."""
        from docwen_plugin_proofread import ProofreadPlugin

        plugin = ProofreadPlugin()
        assert plugin.can_handle("pdf", "docx", "validate") is False
        assert plugin.can_handle("txt", "docx", "validate") is False
        assert plugin.can_handle("png", "docx", "validate") is False

    def test_manifest_has_three_routes(self) -> None:
        """The manifest declares two direct parsers plus one composed category route."""
        from docwen_plugin_proofread import ProofreadPlugin

        plugin = ProofreadPlugin()
        assert len(plugin.manifest.routes) == 3

    def test_manifest_declares_only_consumed_proofread_action_options(self) -> None:
        """Proofread action routes expose the plugin request keys, not GUI/config internals."""
        from docwen_plugin_proofread import ProofreadPlugin

        plugin = ProofreadPlugin()
        expected = {
            "enable_symbol_pairing",
            "enable_symbol_correction",
            "enable_typos_rule",
            "enable_sensitive_word",
            "skip_code_blocks",
            "skip_quote_blocks",
        }

        routes = {
            (route.source_format, route.target_format, route.action_name): route for route in plugin.manifest.routes
        }
        assert set(routes) == {
            ("docx", "docx", "validate"),
            ("document", "docx", "validate"),
            ("markdown", "markdown", "validate"),
        }
        for route in routes.values():
            properties = route.options_schema["properties"]
            assert set(properties) == expected
            assert all(spec["type"] == "boolean" for spec in properties.values())
            assert properties["enable_symbol_pairing"]["default"] is True
            assert properties["enable_symbol_correction"]["default"] is True
            assert properties["enable_typos_rule"]["default"] is True
            assert properties["enable_sensitive_word"]["default"] is True
            assert properties["skip_code_blocks"]["default"] is True
            assert properties["skip_quote_blocks"]["default"] is False
            assert "symbol_pairing" not in properties
            assert "proofread_rules" not in properties


# ═══════════════════════════════════════════════════════════════════════════
# Pre-conversion moved to application layer — plugin no longer does it
# ═══════════════════════════════════════════════════════════════════════════


@pytest.mark.contract
class TestConcreteValidationBoundary:
    """Verify non-DOCX inputs remain outside the plugin's direct parser boundary."""

    def _create_fake_docx(self, path: str) -> None:
        """Create a valid but empty DOCX so the validator doesn't crash."""
        _create_minimal_docx(path, ["Pre-converted content."])

    def test_non_docx_input_rejected_by_plugin(self) -> None:
        """Direct plugin calls report an unsupported route for non-DOCX input."""
        from docwen_plugin_proofread.plugin import ProofreadPlugin

        with tempfile.TemporaryDirectory() as staging:
            doc_path = os.path.join(staging, "report.doc")
            Path(doc_path).write_bytes(b"fake doc content")

            context = _build_fake_context(
                doc_path,
                staging,
                source_format="doc",
            )
            result = ProofreadPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "unsupported_route"

    def test_docx_still_works_directly(self) -> None:
        """Normal DOCX validation should still work without pre-conversion."""
        from docwen_plugin_proofread.plugin import ProofreadPlugin

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            self._create_fake_docx(docx_path)

            context = _build_fake_context(
                docx_path,
                staging,
                source_format="docx",
            )
            result = ProofreadPlugin().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1

    def test_non_docx_direct_to_validator_rejected(self) -> None:
        """Passing a non-DOCX file directly to DocxValidator should
        produce an invalid_input error."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            doc_path = os.path.join(staging, "report.doc")
            Path(doc_path).write_bytes(b"fake doc content")

            context = _build_fake_context(
                doc_path,
                staging,
                source_format="doc",
            )
            result = DocxValidator().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "invalid_input"


# ═══════════════════════════════════════════════════════════════════════════
# Common module — only non-preconversion helpers remain
# ═══════════════════════════════════════════════════════════════════════════


@pytest.mark.unit
class TestCommonHelpers:
    """Unit tests for remaining _common.py helpers."""

    def test_new_artifact_id_format(self) -> None:
        """new_artifact_id should return a correctly formatted string."""
        from docwen_plugin_proofread._common import new_artifact_id

        aid = new_artifact_id()
        assert aid.startswith("proofread-")
        assert len(aid) == len("proofread-") + 12  # uuid hex (12 chars)

    def test_file_size_existing(self, tmp_path) -> None:
        """file_size should return the size of an existing file."""
        from docwen_plugin_proofread._common import file_size

        f = tmp_path / "test.txt"
        f.write_text("hello")
        assert file_size(str(f)) == 5

    def test_file_size_missing(self) -> None:
        """file_size should return 0 for a non-existent file."""
        from docwen_plugin_proofread._common import file_size

        assert file_size("/nonexistent/path") == 0

    def test_request_source_format_ignores_filename_suffix(self, tmp_path: Path) -> None:
        """Parser selection must consume FileRef.format, not the path suffix."""
        from docwen_plugin_proofread._common import request_source_format

        path = tmp_path / "misleading.md"
        path.write_bytes(b"placeholder")
        context = _build_fake_context(str(path), str(tmp_path), source_format="docx")

        assert request_source_format(context) == "docx"


# ═══════════════════════════════════════════════════════════════════════════
# Regression: existing DOCX/MD routes are unchanged
# ═══════════════════════════════════════════════════════════════════════════


class TestNoRegression:
    """Application preconversion must not break direct DOCX and Markdown routes."""

    @pytest.mark.contract
    def test_docx_validation_still_works(self) -> None:
        """Normal DOCX→DOCX validate must still succeed."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_minimal_docx(docx_path, ["Hello world."])

            context = _build_fake_context(
                docx_path,
                staging,
                source_format="docx",
            )
            result = DocxValidator().convert(context)
            assert result.success is True

    @pytest.mark.integration
    def test_runtime_docx_validation_uses_admitted_format_despite_txt_suffix(self, tmp_path: Path) -> None:
        """A post-admission DOCX keeps its parser when the visible suffix is TXT."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_proofread import ProofreadPlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        input_path = tmp_path / "admitted-docx.txt"
        _create_minimal_docx(str(input_path), ["Hello world."])
        output_dir = tmp_path / "out"
        output_dir.mkdir()
        workspace_root = tmp_path / "workspace"

        registry = PluginRegistry()
        registry.register(ProofreadPlugin())
        task_manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id="proofread-admitted-docx-wrong-suffix",
            input_refs=[
                FileRef(
                    path=str(input_path),
                    format="docx",
                    category="document",
                    size_bytes=input_path.stat().st_size,
                )
            ],
            target_format="docx",
            action_name="validate",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_manager.execute_single(request)

        assert result.success is True, f"unexpected error: {result.error}"
        assert result.artifacts[0].metadata["source_format"] == "docx"
        assert Path(result.artifacts[0].staging_path).suffix == ".docx"

    @pytest.mark.contract
    def test_markdown_validation_uses_admitted_format_despite_docx_suffix(self, tmp_path: Path) -> None:
        """Markdown text is parsed as text while retaining its display filename."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        input_path = tmp_path / "admitted-markdown.docx"
        input_path.write_text("Clean markdown.", encoding="utf-8")
        context = _build_fake_context(
            str(input_path),
            str(tmp_path),
            target_format="markdown",
            action_name="validate",
            source_format="markdown",
        )

        result = MarkdownValidator().convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        artifact = result.artifacts[0]
        assert artifact.metadata["source_format"] == "markdown"
        report = json.loads(Path(artifact.staging_path).read_text(encoding="utf-8"))
        assert report["file"] == "admitted-markdown.docx"

    @pytest.mark.contract
    def test_md_validation_still_works(self) -> None:
        """Normal MD→MD validate must still succeed."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text("Clean markdown.", encoding="utf-8")

            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)
            assert result.success is True

    @pytest.mark.contract
    def test_cancellation_still_works(self) -> None:
        """Cancellation must still work for DOCX validation."""
        from docwen_core.errors import CancellationRequested
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_minimal_docx(docx_path, ["Test."])

            context = _build_fake_context(
                docx_path,
                staging,
                source_format="docx",
                pre_cancelled=True,
            )
            with pytest.raises(CancellationRequested):
                DocxValidator().convert(context)
