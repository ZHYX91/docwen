"""Focused tests split from test_proofread_plugin.py."""

from __future__ import annotations

from ._proofread_plugin_support import (
    Path,
    _build_fake_context,
    _count_by_key,
    _create_test_docx,
    _load_proofread_old_system_fixture,
    _proofread_options_from_fixture,
    _proofread_rules_from_fixture,
    os,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.golden


@pytest.mark.contract
class TestDocxValidation:
    def test_docx_validation_matches_old_system_semantic_fixture(self) -> None:
        """DOCX validation should surface the old systems' normalized proofread issues as comments."""
        from docwen_plugin_proofread.anchor_report import _extract_comments
        from docwen_plugin_proofread.docx_validator import DocxValidator

        fixture = _load_proofread_old_system_fixture()
        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "old_system_probe.docx")
            _create_test_docx(docx_path, [fixture["input_text"]])
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
                options=_proofread_options_from_fixture(fixture),
                proofread_rules=_proofread_rules_from_fixture(fixture),
            )

            result = DocxValidator().convert(context)

            assert result.success is True
            assert result.artifacts[0].metadata["errors_found"] == len(fixture["expected_issues"])
            comments = _extract_comments(Path(result.artifacts[0].staging_path))
            assert len(comments) == len(fixture["expected_issues"])
            assert _count_by_key(comment.author.removeprefix("DocWen-") for comment in comments) == {
                "sensitive": 1,
                "typo": 1,
                "symbol": 3,
                "pairing": 1,
            }
            comment_texts = [comment.text for comment in comments]
            for expected in fixture["expected_issues"]:
                assert any(expected["error_text"] in text for text in comment_texts)

    def test_docx_no_errors(self) -> None:
        """A clean DOCX with correct text should produce zero errors."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_test_docx(
                docx_path,
                [
                    "这是一个正确的句子。",
                    "第二段也没有问题。",
                ],
            )
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
            )
            result = DocxValidator().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.suggested_name.endswith("_checked.docx")
            assert artifact.is_primary is True
            assert os.path.isfile(artifact.staging_path)

            # Re-open to verify it's valid DOCX
            from docx import Document

            doc = Document(artifact.staging_path)
            assert len(doc.paragraphs) >= 2

    def test_docx_with_typos(self) -> None:
        """DOCX validation pipeline runs with typos rule enabled (empty default map)."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_test_docx(
                docx_path,
                [
                    "这是一个标准测试段落。",
                ],
            )
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
                options={
                    "enable_typos_rule": True,
                    "enable_symbol_pairing": False,
                    "enable_symbol_correction": False,
                    "enable_sensitive_word": False,
                },
            )
            result = DocxValidator().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            # Default typo map is empty, so no errors expected
            assert result.artifacts[0].metadata["errors_found"] == 0

    def test_docx_with_symbol_pairing_issues(self) -> None:
        """Unmatched brackets should be detected."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_test_docx(
                docx_path,
                [
                    "这是一个（不完整的括号测试。",  # missing closing ）
                ],
            )
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
                options={
                    "enable_symbol_pairing": True,
                    "enable_typos_rule": False,
                    "enable_symbol_correction": False,
                    "enable_sensitive_word": False,
                },
            )
            result = DocxValidator().convert(context)

            assert result.success is True
            assert result.artifacts[0].metadata["errors_found"] >= 1

    def test_docx_balanced_symmetric_quotes_do_not_create_pairing_issues(self) -> None:
        """The real DOCX route consumes the corrected symmetric-quote semantics."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "balanced-quotes.docx")
            _create_test_docx(docx_path, ['He said "ok". It\'s ready.'])
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
                options={
                    "enable_symbol_pairing": True,
                    "enable_typos_rule": False,
                    "enable_symbol_correction": False,
                    "enable_sensitive_word": False,
                },
            )

            result = DocxValidator().convert(context)

            assert result.success is True
            assert result.artifacts[0].metadata["errors_found"] == 0

    def test_docx_with_fullwidth_digits(self) -> None:
        """Fullwidth digits should be flagged by symbol correction."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_test_docx(
                docx_path,
                [
                    "价格是１２３元。",  # fullwidth digits
                ],
            )
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
                options={
                    "enable_symbol_correction": True,
                    "enable_symbol_pairing": False,
                    "enable_typos_rule": False,
                    "enable_sensitive_word": False,
                },
            )
            result = DocxValidator().convert(context)

            assert result.success is True
            assert result.artifacts[0].metadata["errors_found"] >= 3  # 3 fullwidth digits

    def test_docx_all_checks_disabled(self) -> None:
        """When all checks are disabled, should skip gracefully."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_test_docx(docx_path, ["Some text."])
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
                options={
                    "enable_symbol_pairing": False,
                    "enable_symbol_correction": False,
                    "enable_typos_rule": False,
                    "enable_sensitive_word": False,
                },
            )
            result = DocxValidator().convert(context)

            assert result.success is True
            assert any(d.code == "PROOFREAD-SKIPPED" for d in result.diagnostics)

    def test_docx_empty_paragraphs(self) -> None:
        """A DOCX with empty paragraphs should not crash."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_test_docx(
                docx_path,
                [
                    "",  # empty
                    "   ",  # whitespace only
                    "Text with content.",
                ],
            )
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
            )
            result = DocxValidator().convert(context)

            assert result.success is True

    def test_docx_corrupted_file(self) -> None:
        """A non-DOCX file should produce an error."""
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            bad_path = os.path.join(staging, "bad.docx")
            Path(bad_path).write_text("This is not a valid DOCX file")

            context = _build_fake_context(
                bad_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
            )
            result = DocxValidator().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.diagnostic_code in (
                "PROOFREAD-ERROR",
                "PROOFREAD-CORRUPTED-DOCX",
            )

    def test_docx_cancellation(self) -> None:
        """A pre-cancelled context should raise before any work."""
        from docwen_core.errors import CancellationRequested
        from docwen_plugin_proofread.docx_validator import DocxValidator

        with tempfile.TemporaryDirectory() as staging:
            docx_path = os.path.join(staging, "test.docx")
            _create_test_docx(docx_path, ["Test."])
            context = _build_fake_context(
                docx_path,
                staging,
                target_format="document",
                action_name="validate",
                source_format="docx",
                pre_cancelled=True,
            )
            with pytest.raises(CancellationRequested):
                DocxValidator().convert(context)
