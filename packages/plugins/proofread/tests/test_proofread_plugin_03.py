"""Focused tests split from test_proofread_plugin.py."""

from __future__ import annotations

from ._proofread_plugin_support import (
    Path,
    _build_fake_context,
    _create_test_docx,
    _expected_markdown_issues_from_fixture,
    _load_proofread_old_system_fixture,
    _markdown_issue_projection,
    _proofread_options_from_fixture,
    _proofread_rules_from_fixture,
    json,
    os,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.golden


class TestMarkdownValidation:
    @pytest.mark.contract
    def test_md_validation_matches_old_system_semantic_fixture(self) -> None:
        """Markdown validation should preserve old-system issue positions after report mapping."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        fixture = _load_proofread_old_system_fixture()
        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "old_system_probe.md")
            Path(md_path).write_text(fixture["input_text"] + "\n", encoding="utf-8")
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
                options=_proofread_options_from_fixture(fixture),
                proofread_rules=_proofread_rules_from_fixture(fixture),
            )

            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            assert [_markdown_issue_projection(issue) for issue in report["issues"]] == (
                _expected_markdown_issues_from_fixture(fixture)
            )
            assert report["summary"] == fixture["expected_summary"]

    @pytest.mark.integration
    def test_md_validation_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path) -> None:
        """Markdown proofread JSON report should be finalized into the output dir."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_proofread.plugin import ProofreadPlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture = _load_proofread_old_system_fixture()
        input_file = tmp_path / "old_system_probe.md"
        input_file.write_text(fixture["input_text"] + "\n", encoding="utf-8")
        output_dir = tmp_path / "out"
        output_dir.mkdir()
        workspace_root = tmp_path / "workspace"
        registry = PluginRegistry()
        registry.register(ProofreadPlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
            proofread_rules=_proofread_rules_from_fixture(fixture),
        )
        request = ConversionRequest(
            request_id="proofread-md-finalizer-old-system-fixture",
            input_refs=[
                FileRef(
                    path=str(input_file),
                    format="markdown",
                    category="text",
                    size_bytes=input_file.stat().st_size,
                )
            ],
            target_format="markdown",
            action_name="validate",
            options=_proofread_options_from_fixture(fixture),
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_mgr.execute_single(request)

        assert result.success, f"unexpected error: {result.error}"
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        report_path = Path(artifact.staging_path)
        assert report_path.parent == output_dir
        assert report_path.name == "old_system_probe_proofread_report.json"
        assert artifact.media_type == "application/json"
        assert artifact.metadata["source_format"] == "markdown"
        assert artifact.metadata["issues_found"] == len(fixture["expected_issues"])
        expected_checks = fixture["rules"]["checks_enabled"]
        assert artifact.metadata["checks_enabled"] == expected_checks
        assert any(d.code == "PROOFREAD-OK" for d in result.diagnostics)
        report = json.loads(report_path.read_text(encoding="utf-8"))
        assert [_markdown_issue_projection(issue) for issue in report["issues"]] == (
            _expected_markdown_issues_from_fixture(fixture)
        )
        assert report["summary"] == fixture["expected_summary"]
        assert report["checks_enabled"] == expected_checks
        assert str(workspace_root) not in report_path.read_text(encoding="utf-8")

    @pytest.mark.contract
    def test_md_no_errors(self) -> None:
        """A clean Markdown file should produce an empty report."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text(
                "This is a clean Markdown file.\n\nNo issues here.\n",
                encoding="utf-8",
            )
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.media_type == "application/json"
            assert os.path.isfile(artifact.staging_path)

            report = json.loads(Path(artifact.staging_path).read_text(encoding="utf-8"))
            assert report["schema"] == "docwen.proofread_report.v2"
            assert len(report["issues"]) == 0

    @pytest.mark.contract
    def test_md_with_symbol_issues(self) -> None:
        """A Markdown file with symbol/fullwidth issues should produce a report."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text(
                "价格是１２３元。\n",  # fullwidth digits
                encoding="utf-8",
            )
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            assert len(report["issues"]) >= 1
            assert any(i["rule_key"] == "symbol_correct" for i in report["issues"])

    @pytest.mark.contract
    def test_md_uses_injected_proofread_rules(self) -> None:
        """Injected proofread rules must be consumed instead of path-based loading."""
        from docwen_core.models.proofread import ProofreadRules
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text("我己经完成。\n", encoding="utf-8")
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
                options={
                    "enable_symbol_pairing": False,
                    "enable_symbol_correction": False,
                    "enable_typos_rule": True,
                    "enable_sensitive_word": False,
                },
                proofread_rules=ProofreadRules(typos_map={"已": ("己",)}),
            )

            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            assert any(issue["rule_key"] == "typo" for issue in report["issues"])

    @pytest.mark.contract
    def test_md_with_symbol_pairing(self) -> None:
        """Unmatched brackets in Markdown should be reported."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text(
                "这段文字（括号没闭合。\n",
                encoding="utf-8",
            )
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            pairing_issues = [i for i in report["issues"] if i["rule_key"] == "symbol_pair"]
            assert len(pairing_issues) >= 1

    @pytest.mark.contract
    def test_md_balanced_and_escaped_quotes_do_not_create_pairing_issues(self) -> None:
        """The real Markdown route honors quote toggling and Markdown escapes."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "balanced-quotes.md")
            Path(md_path).write_text(
                'He said "ok".\nEscaped \\"literal.\nIt\'s ready.\n',
                encoding="utf-8",
            )
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
                options={
                    "enable_symbol_pairing": True,
                    "enable_typos_rule": False,
                    "enable_symbol_correction": False,
                    "enable_sensitive_word": False,
                },
            )

            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            assert [issue for issue in report["issues"] if issue["rule_key"] == "symbol_pair"] == []

    @pytest.mark.contract
    def test_md_code_blocks_sanitized(self) -> None:
        """Code blocks should be sanitized — no issues reported inside them."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text(
                "Before code.\n\n```\n我己经在代码块里（这个不应该被检查\n```\n\nAfter code.\n",
                encoding="utf-8",
            )
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            # No issues should come from inside the code block
            # (the typos/brackets inside ```...``` should be ignored)
            # There might be some issues in "Before code" or "After code",
            # but "己" and "（" inside the code block should not appear
            for issue in report["issues"]:
                # The code block lines should not contain issues
                assert "代码块" not in issue.get("error_text", "")

    @pytest.mark.contract
    def test_md_yaml_frontmatter_sanitized(self) -> None:
        """YAML frontmatter should be sanitized."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text(
                "---\n"
                "title: 测试（文档\n"  # unmatched bracket in YAML
                "date: 2024-01-01\n"
                "---\n\n"
                "实际内容在这里。\n",
                encoding="utf-8",
            )
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            # The unmatched bracket inside YAML frontmatter should be ignored
            pairing_issues = [i for i in report["issues"] if i["rule_key"] == "symbol_pair"]
            # No pairing issues because the frontmatter bracket is sanitized
            assert len(pairing_issues) == 0, f"YAML frontmatter should be sanitized, got: {pairing_issues}"

    @pytest.mark.contract
    def test_md_all_checks_disabled(self) -> None:
        """Disabled checks still produce a successful, attributable empty report."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text("Some text.", encoding="utf-8")
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
                options={
                    "enable_symbol_pairing": False,
                    "enable_symbol_correction": False,
                    "enable_typos_rule": False,
                    "enable_sensitive_word": False,
                },
            )
            result = MarkdownValidator().convert(context)

            assert result.success is True
            assert any(d.code == "PROOFREAD-SKIPPED" for d in result.diagnostics)
            assert len(result.artifacts) == 1
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            assert report["schema"] == "docwen.proofread_report.v2"
            assert report["issues"] == []
            assert report["summary"] == {}

    @pytest.mark.contract
    def test_md_empty_file(self) -> None:
        """An empty Markdown file should produce zero issues."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text("", encoding="utf-8")
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            assert len(report["issues"]) == 0

    @pytest.mark.contract
    def test_md_missing_file(self) -> None:
        """A non-existent file should produce an error."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            bad_path = os.path.join(staging, "nonexistent.md")
            context = _build_fake_context(
                bad_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.diagnostic_code == "PROOFREAD-INVALID-INPUT"

    @pytest.mark.contract
    def test_md_report_summary(self) -> None:
        """The report should include a summary grouped by rule_key."""
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text(
                "我己经完成，１２３元（未闭合。\n",
                encoding="utf-8",
            )
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
            )
            result = MarkdownValidator().convert(context)

            assert result.success is True
            report = json.loads(Path(result.artifacts[0].staging_path).read_text(encoding="utf-8"))
            assert "summary" in report
            assert isinstance(report["summary"], dict)
            # Default typo map is empty, so no typo errors;
            # but fullwidth digits and unmatched bracket should be caught.
            assert report["summary"].get("symbol_correct", 0) >= 1
            assert report["summary"].get("symbol_pair", 0) >= 1

    @pytest.mark.contract
    def test_md_cancellation(self) -> None:
        """A pre-cancelled context should raise before any work."""
        from docwen_core.errors import CancellationRequested
        from docwen_plugin_proofread.md_validator import MarkdownValidator

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text("Test", encoding="utf-8")
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
                pre_cancelled=True,
            )
            with pytest.raises(CancellationRequested):
                MarkdownValidator().convert(context)


@pytest.mark.contract
class TestPluginDispatch:
    def test_plugin_routes_docx_validate(self) -> None:
        """The ProofreadPlugin should route concrete docx validation."""
        from docwen_plugin_proofread import ProofreadPlugin

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_test_docx(docx_path, ["Hello."])
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="docx",
                action_name="validate",
                source_format="docx",
                options={
                    "enable_typos_rule": False,
                    "enable_symbol_correction": False,
                    "enable_symbol_pairing": False,
                    "enable_sensitive_word": False,
                },
            )
            result = ProofreadPlugin().convert(context)

            assert result.success is True
            assert any(d.code == "PROOFREAD-SKIPPED" for d in result.diagnostics)

    def test_plugin_routes_md_validate(self) -> None:
        """The ProofreadPlugin should route validate+markdown to MarkdownValidator."""
        from docwen_plugin_proofread import ProofreadPlugin

        with tempfile.TemporaryDirectory() as staging:
            md_path = os.path.join(staging, "test.md")
            Path(md_path).write_text("Hello.", encoding="utf-8")
            context = _build_fake_context(
                md_path,
                staging,
                target_format="markdown",
                action_name="validate",
                source_format="markdown",
                options={
                    "enable_typos_rule": False,
                    "enable_symbol_correction": False,
                    "enable_symbol_pairing": False,
                    "enable_sensitive_word": False,
                },
            )
            result = ProofreadPlugin().convert(context)

            assert result.success is True
            assert any(d.code == "PROOFREAD-SKIPPED" for d in result.diagnostics)

    def test_plugin_can_handle(self) -> None:
        """can_handle should match expected routes."""
        from docwen_plugin_proofread import ProofreadPlugin

        plugin = ProofreadPlugin()
        assert plugin.can_handle("docx", "docx", "validate") is True
        assert plugin.can_handle("markdown", "markdown", "validate") is True
        assert plugin.can_handle("pdf", "md", "") is False
        assert plugin.can_handle("docx", "docx", "invalid") is False

    def test_plugin_fallback_invalid_action(self) -> None:
        """Unknown action should return a typed unsupported-route failure."""
        from docwen_plugin_proofread import ProofreadPlugin

        with tempfile.TemporaryDirectory() as staging:
            dummy = os.path.join(staging, "dummy.txt")
            Path(dummy).write_text("test")
            context = _build_fake_context(
                dummy,
                staging,
                target_format="pdf",
                action_name="convert",
                source_format="document",
            )
            result = ProofreadPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "unsupported_route"
            assert result.error.diagnostic_code == "PROOFREAD-UNSUPPORTED-ROUTE"
