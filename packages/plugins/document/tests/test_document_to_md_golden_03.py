"""Focused tests split from test_document_to_md_golden.py."""

from __future__ import annotations

from ._document_to_md_golden_support import (
    ConversionRequest,
    OutputPolicy,
    Path,
    _build_runtime_pipeline,
    _run_conversion,
    os,
    pytest,
    re,
)

pytestmark = [pytest.mark.golden, pytest.mark.contract]


class TestDocxToMdGolden:
    """Golden parity tests for ROUTE-DOC-001: document → md.

    Uses a programmatically-generated sample DOCX with known content
    (headings, paragraphs with formatting, tables) for structural
    validation, plus a real template for integration smoke testing.
    """

    @pytest.fixture
    def pipeline(self):
        """Build the full runtime pipeline with the real DocumentPlugin."""
        plugin, task_mgr, ws_mgr, ws_root = _build_runtime_pipeline()
        yield plugin, task_mgr, ws_mgr
        ws_mgr.cleanup_all()
        import shutil

        shutil.rmtree(ws_root, ignore_errors=True)

    @staticmethod
    def _markdown_table_cells(line: str) -> list[str]:
        row = line.strip()
        if row.startswith("|"):
            row = row[1:]
        if row.endswith("|"):
            row = row[:-1]
        return [cell.replace(r"\|", "|").strip() for cell in re.split(r"(?<!\\)\|", row)]

    @staticmethod
    def _verify_tables_well_formed(content: str) -> None:
        """Verify that Markdown tables in content are well-formed and have data rows."""
        lines = content.splitlines()
        in_table = False
        separator_seen = False
        data_rows_seen = 0
        column_count = 0

        for line in lines:
            stripped = line.strip()
            is_table_line = stripped.startswith("|") and stripped.endswith("|")

            if is_table_line:
                if not in_table:
                    in_table = True
                    separator_seen = False
                    data_rows_seen = 0
                    column_count = len(TestDocxToMdGolden._markdown_table_cells(stripped))
                else:
                    if not separator_seen:
                        separator_seen = True
                        parts = TestDocxToMdGolden._markdown_table_cells(stripped)
                        assert all(set(p) <= {"-", ":", " "} and "-" in p for p in parts), (
                            f"Malformed table separator: {stripped}"
                        )
                    else:
                        cols = len(TestDocxToMdGolden._markdown_table_cells(stripped))
                        assert cols == column_count, f"Table column count mismatch: expected {column_count}, got {cols}"
                        data_rows_seen += 1
            else:
                if in_table and stripped == "":
                    in_table = False
                    separator_seen = False
                    assert data_rows_seen >= 1, f"Table should have at least 1 data row, got {data_rows_seen}"

    def test_reserved_options_emit_warning_diagnostics(self, pipeline, sample_docx_path, tmp_path) -> None:
        """remove_numbering is now implemented — no reserved warning expected.

        ``to_md_keep_images`` is functional (DFG-003).
        ``remove_numbering`` is functional (gongwen-shared-infra Task 10).
        """
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_reserved"
        output_dir.mkdir()

        result = _run_conversion(task_mgr, sample_docx_path, output_dir)

        assert result.success
        # No reserved-option warnings expected (both to_md_keep_images and
        # remove_numbering are now implemented).
        warnings = [d for d in result.diagnostics if d.level == "warning"]
        reserved = [d for d in warnings if "reserved" in d.message.lower() or d.code == "DOCX2MD-RESERVED-OPTION"]
        assert len(reserved) == 0, (
            f"Expected 0 reserved-option warnings, got {len(reserved)}. "
            f"Diagnostics: {[(d.level, d.code, d.message[:60]) for d in result.diagnostics]}"
        )

    def test_removed_optimize_for_type_option_is_rejected(self, pipeline, sample_docx_path, tmp_path) -> None:
        """Removed route options fail closed before plugin execution."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_optimize_legacy"
        output_dir.mkdir()

        result = _run_conversion(
            task_mgr,
            sample_docx_path,
            output_dir,
            optimize_for_type="invoice_cn",
        )

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "invalid_input"
        assert result.error.diagnostic_code == "ROUTE_OPTIONS_UNSUPPORTED"
        assert "optimize_for_type" in result.error.message

    def test_plugin_does_not_depend_on_runtime(self, pipeline, sample_docx_path, tmp_path) -> None:
        """The DocumentPlugin must not import runtime internals."""
        import sys

        plugin_pkg = "docwen_plugin_document"
        forbidden_prefixes = [
            "docwen_runtime.",
            "docwen_application.",
            "docwen_gui.",
            "docwen_cli.",
            "docwen_bundle.",
        ]

        relevant_modules = {k: v for k, v in sys.modules.items() if k.startswith(plugin_pkg)}

        for mod_name, mod in relevant_modules.items():
            for attr_name, attr_val in mod.__dict__.items():
                attr_mod = getattr(attr_val, "__module__", "")
                if attr_mod and attr_mod.startswith(plugin_pkg):
                    continue
                if attr_mod:
                    for forbidden in forbidden_prefixes:
                        assert not attr_mod.startswith(forbidden), (
                            f"Plugin module {mod_name} imports {attr_mod} "
                            f"({attr_name}), which is forbidden ({forbidden})"
                        )

    def test_plugin_only_writes_to_staging(self, pipeline, sample_docx_path, tmp_path) -> None:
        """GOLDEN-002: Plugin writes to staging; finalizer places output."""
        _plugin, task_mgr, _ws_mgr = pipeline

        final_dir = tmp_path / "final_output"
        final_dir.mkdir()

        result = _run_conversion(task_mgr, sample_docx_path, final_dir)
        assert result.success

        final_path = result.artifacts[0].staging_path
        assert final_path.startswith(str(final_dir)), f"Final artifact not in output dir: {final_path}"
        assert os.path.isfile(final_path)


class TestDocumentPluginDirect:
    """Test the plugin directly, without the full runtime pipeline."""

    def test_plugin_convert_directly_with_fake_context(self, sample_docx_path, tmp_path) -> None:
        """The DocumentPlugin can accept a fake context and produce a result."""
        from tests.support.config import FakeConfigView
        from tests.support.execution import FakeExecutionContext
        from tests.support.logging import FakePluginLogger
        from tests.support.progress import FakeProgressSink
        from tests.support.workspace import FakeWorkspaceHandle

        from docwen_core.cancellation import CancellationToken
        from docwen_core.models.file_ref import FileRef
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()

        staging_dir = tmp_path / "staging"
        staging_dir.mkdir()

        workspace = FakeWorkspaceHandle(str(sample_docx_path), str(staging_dir))
        progress = FakeProgressSink()
        logger = FakePluginLogger()
        config = FakeConfigView()

        ctx = FakeExecutionContext(
            request=ConversionRequest(
                request_id="direct-test-001",
                input_refs=[
                    FileRef(
                        path=str(sample_docx_path),
                        format="docx",
                        category="document",
                    )
                ],
                target_format="md",
                options={"to_md_keep_images": True, "remove_numbering": True},
                output_policy=OutputPolicy(),
            ),
            workspace=workspace,
            config=config,
            progress=progress,
            cancellation=CancellationToken().view(),
            logger=logger,
            numbering_registry=None,
            proofread_rules=None,
        )
        result = plugin.convert(ctx)

        assert result.success, f"Direct conversion failed: {result.error.message if result.error else 'unknown'}"
        assert len(result.artifacts) == 1
        assert result.artifacts[0].is_primary is True

        # Verify staging file exists and contains markdown
        staging_file = result.artifacts[0].staging_path
        assert os.path.isfile(staging_file), f"Staging file not created: {staging_file}"
        content = Path(staging_file).read_text(encoding="utf-8")
        assert len(content) > 0
        assert "# " in content, f"Output should contain Markdown headings. Content:\n{content[:200]}"

        # Progress should have been reported
        assert len(progress.events) >= 2
