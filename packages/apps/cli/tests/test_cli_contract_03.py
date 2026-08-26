"""Focused tests split from test_cli_contract.py."""

from __future__ import annotations

from ._cli_contract_support import (
    ExitCode,
    MagicMock,
    _make_execution_args,
    pytest,
)
from ._cli_contract_support import (
    _runtime_route_contract as _runtime_route_contract,
)

pytestmark = pytest.mark.unit


class TestExecuteConvertActionPath:
    """All execution paths use the same resolve_cli_action helper."""

    def test_plain_txt_request_preserves_txt_format_with_markdown_workflow_fallback(self, tmp_path) -> None:
        """Plain TXT remains TXT even though it shares the Markdown workflow."""
        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "plain.txt"
        src.write_text("ordinary text without Markdown syntax\nsecond line\n", encoding="utf-8")
        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="", to="docx", files=[str(src)])
        exit_code = execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert exit_code == 0
        assert request.input_refs[0].format == "txt"
        assert request.input_refs[0].category == "markdown"

    def test_plain_txt_dry_run_exposes_markdown_category_as_deferred_fallback(self, tmp_path, capsys) -> None:
        """TXT dry-run exposes candidates without pretending fallback already won."""
        import json

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "plain.txt"
        src.write_text("ordinary text without Markdown syntax\nsecond line\n", encoding="utf-8")
        args = _make_execution_args(json=True, dry_run=True, to="docx", files=[str(src)])

        exit_code = execute_convert(args, controller=MagicMock(has_runtime=True))

        assert exit_code == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload["data"]["detected_format"] == "txt"
        assert payload["data"]["workflow_category"] == "markdown"
        assert payload["data"]["routing"] == {
            "status": "deferred_to_runtime",
            "source_format": "txt",
            "source_category": "markdown",
            "source_candidates": ["txt", "markdown"],
            "category_fallback": "markdown",
            "target_format": "docx",
            "action_name": "",
        }

    def test_plain_txt_human_dry_run_calls_out_deferred_resolution(self, tmp_path, capsys) -> None:
        """Human output distinguishes candidates from an actually resolved route."""
        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "plain.txt"
        src.write_text("ordinary text without Markdown syntax\n", encoding="utf-8")
        args = _make_execution_args(dry_run=True, to="docx", files=[str(src)])

        exit_code = execute_convert(args, controller=MagicMock(has_runtime=True))

        assert exit_code == 0
        output = capsys.readouterr().out
        assert "路由候选（按顺序）: txt -> markdown" in output
        assert "类别回退: markdown" in output
        assert "路由解析: 执行时由 runtime 确定" in output

    def test_template_uses_detected_markdown_when_docx_suffix_is_explicitly_accepted(self, tmp_path) -> None:
        """Content-first admission also drives template option eligibility."""
        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "actually-markdown.docx"
        src.write_text("# Title\n\ncontent\n", encoding="utf-8")
        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)
        args = _make_execution_args(
            action="",
            to="docx",
            files=[str(src)],
            template=f"template.docx.{'a' * 64}",
            use_detected_format=True,
            quiet=True,
        )

        exit_code = execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert exit_code == 0
        assert request.input_refs[0].format == "markdown"
        assert request.input_refs[0].category == "markdown"
        assert request.options["template_name"] == f"template.docx.{'a' * 64}"

    def test_xlsx_content_with_ods_suffix_still_receives_xlsx_only_options(self, tmp_path, monkeypatch) -> None:
        """XLSX→ODS protection options follow content, never the suffix."""
        from openpyxl import Workbook

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "actually-xlsx.ods"
        workbook = Workbook()
        worksheet = workbook.active
        assert worksheet is not None
        worksheet["A1"] = "value"
        workbook.save(str(src))
        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)
        prompts: list[str] = []
        monkeypatch.setattr("getpass.getpass", lambda prompt: prompts.append(prompt) or "secret")
        args = _make_execution_args(
            action="",
            to="ods",
            files=[str(src)],
            spreadsheet_password_prompt=True,
            allow_spreadsheet_protection_loss=True,
            quiet=True,
        )

        exit_code = execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert exit_code == 0
        assert prompts
        assert request.input_refs[0].format == "xlsx"
        assert request.input_refs[0].category == "spreadsheet"
        assert request.options["spreadsheet_password"] == "secret"
        assert request.options["allow_spreadsheet_protection_loss"] is True

    def test_dry_run_image_exposes_concrete_format_and_deferred_category_fallback(self, tmp_path, capsys) -> None:
        """Dry-run must not claim that a category route was already selected."""
        import json
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "scan.png"
        src.write_bytes(b"\x89PNG\r\n\x1a\n")

        args = _make_execution_args(json=True, dry_run=True, to="webp", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            exit_code = execute_convert(args, controller=MagicMock(has_runtime=True))

        assert exit_code == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload["data"]["detected_format"] == "png"
        assert payload["data"]["routing"] == {
            "status": "deferred_to_runtime",
            "source_format": "png",
            "source_category": "image",
            "source_candidates": ["png", "image"],
            "category_fallback": "image",
            "target_format": "webp",
            "action_name": "",
        }

    def test_aggregate_passes_action_to_request(self, tmp_path) -> None:
        """Aggregate path passes resolved action to aggregate_command."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src1 = tmp_path / "a.pdf"
        src1.write_bytes(b"%PDF-1.4\n%\x00\x00\x00\x00\n")
        src2 = tmp_path / "b.pdf"
        src2.write_bytes(b"%PDF-1.4\n%\x00\x00\x00\x00\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_aggregate.return_value = MagicMock(success=True)

        args = _make_execution_args(
            action="merge_pdfs",
            command_path="merge pdf",
            to="",
            files=[str(src1), str(src2)],
        )
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src1), str(src2)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request, action_name = mock_controller.execute_aggregate.call_args.args
        assert action_name == "merge_pdfs"
        assert request.action_name == "merge_pdfs"

    @pytest.mark.parametrize("target", ["md", "pdf", "webp"])
    def test_image_standard_conversion_preserves_concrete_format_and_image_category(
        self, tmp_path, target: str
    ) -> None:
        """Runtime receives both concrete content format and category fallback."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "scan.png"
        src.write_bytes(b"\x89PNG\r\n\x1a\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="", to=target, files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.action_name == ""
        assert request.target_format == target
        assert request.input_refs[0].format == "png"
        assert request.input_refs[0].category == "image"

    @pytest.mark.parametrize(
        ("filename", "content", "expected_format"),
        [
            ("actually-png.jpg", b"\x89PNG\r\n\x1a\n", "png"),
            ("actually-jpeg.png", b"\xff\xd8\xff\xe0", "jpeg"),
        ],
    )
    def test_image_suffix_mismatch_keeps_content_format(
        self,
        tmp_path,
        filename: str,
        content: bytes,
        expected_format: str,
    ) -> None:
        """Same-family suffix mismatches never rewrite the admitted format."""
        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / filename
        src.write_bytes(content)
        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="", to="webp", files=[str(src)], quiet=True)
        exit_code = execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert exit_code == 0
        assert request.input_refs[0].format == expected_format
        assert request.input_refs[0].category == "image"

    def test_layout_render_request_keeps_cli_render_dpi(self, tmp_path) -> None:
        """Layout→image routes consume render_dpi; other no-action routes do not."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "layout.pdf"
        src.write_bytes(b"%PDF-1.4\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(to="png", files=[str(src)], dpi=300)
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.input_refs[0].category == "layout"
        assert request.target_format == "png"
        assert request.options["render_dpi"] == 300

    def test_layout_pdf_normalize_rejects_undeclared_render_dpi(self, tmp_path) -> None:
        """Layout→PDF normalize/passthrough routes do not declare render_dpi."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "layout.pdf"
        src.write_bytes(b"%PDF-1.4\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(to="pdf", files=[str(src)], dpi=300)
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            exit_code = execute_convert(args, controller=mock_controller)

        assert exit_code == int(ExitCode.INVALID_INPUT)
        mock_controller.execute_single.assert_not_called()

    def test_invoice_cn_image_input_uses_image_category_route(self, tmp_path) -> None:
        """invoice_cn has a category route for image invoices but explicit routes for PDF/OFD."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "invoice.jpg"
        src.write_bytes(b"\xff\xd8\xff\xe0")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="invoice_cn", to="md", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.action_name == "invoice_cn"
        assert request.target_format == "md"
        assert request.input_refs[0].format == "jpeg"
        assert request.input_refs[0].category == "image"

    def test_invoice_cn_pdf_input_keeps_explicit_pdf_route(self, tmp_path) -> None:
        """Do not collapse explicit PDF invoice routes to the broader layout category."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "invoice.pdf"
        src.write_bytes(b"%PDF-1.4\n%\x00\x00\x00\x00\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="invoice_cn", to="md", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.action_name == "invoice_cn"
        assert request.input_refs[0].format == "pdf"

    def test_pdf_to_docx_keeps_layout_route_source(self, tmp_path) -> None:
        """PDF→DOCX must reach the layout plugin, where pdf2docx fallback lives."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "layout.pdf"
        src.write_bytes(b"%PDF-1.4\n%\x00\x00\x00\x00\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="", to="docx", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.action_name == ""
        assert request.target_format == "docx"
        assert request.input_refs[0].format == "pdf"
        assert request.input_refs[0].category == "layout"

    @pytest.mark.parametrize("target", ["docx", "wps", "pdf", "xlsx", "csv"])
    def test_txt_input_uses_markdown_source_routes(self, tmp_path, target: str) -> None:
        """Legacy TXT inputs are Markdown sources for document/spreadsheet outputs."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "notes.txt"
        src.write_text("# Title\n\ncontent\n", encoding="utf-8")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="", to=target, files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.action_name == ""
        assert request.target_format == target
        assert request.input_refs[0].format == "markdown"
        assert request.input_refs[0].category == "markdown"

    def test_txt_input_dry_run_uses_markdown_wps_route(self, tmp_path, capsys) -> None:
        """TXT dry-run should stay aligned with Markdown→WPS runtime routing."""
        import json
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "notes.txt"
        src.write_text("# Title\n\ncontent\n", encoding="utf-8")

        args = _make_execution_args(json=True, dry_run=True, to="wps", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            exit_code = execute_convert(args, controller=MagicMock(has_runtime=True))

        assert exit_code == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload["data"]["routing"]["source_candidates"] == ["markdown"]
        assert payload["data"]["routing"]["target_format"] == "wps"

    def test_txt_input_dry_run_uses_markdown_pdf_route(self, tmp_path, capsys) -> None:
        """TXT dry-run should stay aligned with Markdown→PDF runtime routing."""
        import json
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "notes.txt"
        src.write_text("# Title\n\ncontent\n", encoding="utf-8")

        args = _make_execution_args(json=True, dry_run=True, to="pdf", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            exit_code = execute_convert(args, controller=MagicMock(has_runtime=True))

        assert exit_code == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload["data"]["routing"]["source_candidates"] == ["markdown"]
        assert payload["data"]["routing"]["target_format"] == "pdf"

    def test_txt_input_md_numbering_uses_markdown_action_route(self, tmp_path) -> None:
        """TXT inputs can use the Markdown heading-numbering action route."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "notes.txt"
        src.write_text("# Title\n", encoding="utf-8")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="process_md_numbering", to="md", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.action_name == "process_md_numbering"
        assert request.target_format == "md"
        assert request.input_refs[0].format == "markdown"
        assert request.input_refs[0].category == "markdown"
        assert "locale" not in request.options
        assert "yaml_key_labels" not in request.options
