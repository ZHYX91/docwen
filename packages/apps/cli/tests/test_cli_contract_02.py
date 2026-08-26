"""Focused tests split from test_cli_contract.py."""

from __future__ import annotations

import pytest

from ._cli_contract_support import (
    ExitCode,
    MagicMock,
    _make_execution_args,
    _write_ooxml,
)
from ._cli_contract_support import (
    _runtime_route_contract as _runtime_route_contract,
)

pytestmark = pytest.mark.unit


class TestExecuteConvertActionPath:
    """All execution paths use the same resolve_cli_action helper."""

    def test_single_passes_resolved_action_to_request(self, tmp_path) -> None:
        """Single-file path passes resolved action to ConversionRequest."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "test.docx"
        _write_ooxml(src)

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="gongwen", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        call_args = mock_controller.execute_single.call_args[0][0]
        assert call_args.action_name == "gongwen"

    def test_removed_markdown_target_alias_is_not_rewritten(self, tmp_path) -> None:
        """Removed target spellings reach route admission unchanged and fail there."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "test.docx"
        _write_ooxml(src)

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="", to="markdown", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            exit_code = execute_convert(args, controller=mock_controller)

        assert exit_code == int(ExitCode.INVALID_INPUT)
        mock_controller.execute_single.assert_not_called()

    def test_json_batch_keeps_prevalidation_invalid_files_in_results(self, tmp_path, capsys) -> None:
        """JSON batch output should include invalid prevalidation failures without executing them."""
        import json
        from pathlib import Path
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_core.models.result import ConversionResult

        sample = tmp_path / "sample.docx"
        sheet = tmp_path / "sheet.xlsx"
        bad = tmp_path / "notes.bad"
        _write_ooxml(sample)
        _write_ooxml(sheet)
        bad.write_text("unsupported\n", encoding="utf-8")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.side_effect = [
            ConversionResult(task_id="docx", success=True),
            ConversionResult(task_id="xlsx", success=True),
        ]

        args = _make_execution_args(
            json=True,
            batch=True,
            continue_on_error=True,
            files=[str(sample), str(sheet), str(bad)],
        )
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=(
                [str(sample), str(sheet)],
                [(str(bad), "unsupported extension")],
                [],
            ),
        ):
            exit_code = execute_convert(args, controller=mock_controller)

        captured = capsys.readouterr()
        payload = json.loads(captured.out)
        assert exit_code == int(ExitCode.PARTIAL_FAILURE)
        assert mock_controller.execute_single.call_count == 2
        assert captured.err == ""
        assert payload["command"] == "convert"
        assert payload["success"] is False
        assert payload["error"]["code"] == "batch_partial_failure"
        assert payload["data"]["total"] == 3
        assert payload["data"]["succeeded"] == 2
        assert payload["data"]["failed"] == 1
        assert [Path(item["input"]).name for item in payload["data"]["results"]] == [
            "sample.docx",
            "sheet.xlsx",
            "notes.bad",
        ]
        invalid_item = payload["data"]["results"][2]
        assert invalid_item["success"] is False
        assert invalid_item["error"]["code"] == "invalid_input"
        assert invalid_item["error"]["details"] == "unsupported_extension"

    def test_document_to_markdown_request_carries_cli_locale_yaml_labels(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        src = tmp_path / "test.docx"
        _write_ooxml(src)

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        init_cli_locale("de_DE")
        try:
            args = _make_execution_args(action="", to="md", files=[str(src)])
            with patch(
                "docwen_cli.commands.convert.validate_files",
                return_value=([str(src)], [], []),
            ):
                execute_convert(args, controller=mock_controller)
        finally:
            init_cli_locale("zh_CN")

        request = mock_controller.execute_single.call_args[0][0]
        assert request.options["locale"] == "de_DE"
        assert request.options["yaml_key_labels"] == {"title": "Titel", "subtitle": "Untertitel"}
        assert "render_dpi" not in request.options

    def test_document_to_markdown_rejects_undeclared_layout_dpi_option(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "test.docx"
        _write_ooxml(src)

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="", to="md", files=[str(src)], dpi=300)
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            exit_code = execute_convert(args, controller=mock_controller)

        assert exit_code == int(ExitCode.INVALID_INPUT)
        mock_controller.execute_single.assert_not_called()

    def test_spreadsheet_to_markdown_request_carries_cli_locale_yaml_labels(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        src = tmp_path / "book.xlsx"
        _write_ooxml(src)

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        init_cli_locale("de_DE")
        try:
            args = _make_execution_args(action="", to="md", files=[str(src)])
            with patch(
                "docwen_cli.commands.convert.validate_files",
                return_value=([str(src)], [], []),
            ):
                execute_convert(args, controller=mock_controller)
        finally:
            init_cli_locale("zh_CN")

        request = mock_controller.execute_single.call_args[0][0]
        assert request.options["locale"] == "de_DE"
        assert request.options["yaml_key_labels"]["title"] == "Titel"

    def test_image_to_markdown_request_carries_cli_locale_yaml_labels(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        src = tmp_path / "sample.png"
        src.write_bytes(b"\x89PNG\r\n\x1a\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        init_cli_locale("ja_JP")
        try:
            args = _make_execution_args(action="", to="md", files=[str(src)])
            with patch(
                "docwen_cli.commands.convert.validate_files",
                return_value=([str(src)], [], []),
            ):
                execute_convert(args, controller=mock_controller)
        finally:
            init_cli_locale("zh_CN")

        request = mock_controller.execute_single.call_args[0][0]
        assert request.options["locale"] == "ja_JP"
        assert request.options["yaml_key_labels"] == {"title": "タイトル", "subtitle": "サブタイトル"}

    def test_markup_to_markdown_request_carries_cli_locale_yaml_labels(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        src = tmp_path / "page.html"
        src.write_text(
            "<html><head><title>Probe</title></head><body><h1>Probe</h1></body></html>",
            encoding="utf-8",
        )

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        init_cli_locale("de_DE")
        try:
            args = _make_execution_args(action="", to="md", files=[str(src)])
            with patch(
                "docwen_cli.commands.convert.validate_files",
                return_value=([str(src)], [], []),
            ):
                execute_convert(args, controller=mock_controller)
        finally:
            init_cli_locale("zh_CN")

        request = mock_controller.execute_single.call_args[0][0]
        assert request.input_refs[0].format == "html"
        assert request.input_refs[0].category == "markup"
        assert request.options["locale"] == "de_DE"
        assert request.options["yaml_key_labels"]["title"] == "Titel"

    def test_layout_to_markdown_request_carries_cli_locale_yaml_labels(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        src = tmp_path / "layout.pdf"
        src.write_bytes(b"%PDF-1.4\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        init_cli_locale("de_DE")
        try:
            args = _make_execution_args(action="", to="md", files=[str(src)])
            with patch(
                "docwen_cli.commands.convert.validate_files",
                return_value=([str(src)], [], []),
            ):
                execute_convert(args, controller=mock_controller)
        finally:
            init_cli_locale("zh_CN")

        request = mock_controller.execute_single.call_args[0][0]
        assert request.input_refs[0].format == "pdf"
        assert request.input_refs[0].category == "layout"
        assert request.options["locale"] == "de_DE"
        assert request.options["yaml_key_labels"]["title"] == "Titel"

    def test_layout_to_markdown_request_keeps_cli_render_dpi(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "layout.pdf"
        src.write_bytes(b"%PDF-1.4\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(action="", to="md", files=[str(src)], dpi=300)
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.input_refs[0].category == "layout"
        assert request.target_format == "md"
        assert request.options["render_dpi"] == 300

    def test_invoice_cn_to_markdown_request_carries_cli_locale_yaml_labels(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        src = tmp_path / "invoice.pdf"
        src.write_bytes(b"%PDF-1.4\n%\x00\x00\x00\x00\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)
        mock_controller.config_port.get.side_effect = lambda key, default=None: (
            "english" if key == "image.ocr_language" else default
        )

        init_cli_locale("de_DE")
        try:
            args = _make_execution_args(action="invoice_cn", to="md", files=[str(src)])
            with patch(
                "docwen_cli.commands.convert.validate_files",
                return_value=([str(src)], [], []),
            ):
                execute_convert(args, controller=mock_controller)
        finally:
            init_cli_locale("zh_CN")

        request = mock_controller.execute_single.call_args[0][0]
        assert request.action_name == "invoice_cn"
        assert request.target_format == "md"
        assert request.input_refs[0].format == "pdf"
        assert request.options["locale"] == "de_DE"
        assert "ocr_language" not in request.options
        assert request.options["yaml_key_labels"] == {"title": "Titel", "subtitle": "Untertitel"}

    def test_gongwen_to_markdown_request_carries_cli_locale_without_yaml_labels(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        src = tmp_path / "gongwen.docx"
        _write_ooxml(src)

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        init_cli_locale("de_DE")
        try:
            args = _make_execution_args(action="gongwen", to="md", files=[str(src)])
            with patch(
                "docwen_cli.commands.convert.validate_files",
                return_value=([str(src)], [], []),
            ):
                execute_convert(args, controller=mock_controller)
        finally:
            init_cli_locale("zh_CN")

        request = mock_controller.execute_single.call_args[0][0]
        assert request.action_name == "gongwen"
        assert request.target_format == "md"
        assert request.options["locale"] == "de_DE"
        assert "yaml_key_labels" not in request.options

    def test_invoice_cn_to_markdown_request_prefers_cli_ocr_language_over_config(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        src = tmp_path / "invoice.pdf"
        src.write_bytes(b"%PDF-1.4\n%\x00\x00\x00\x00\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)
        mock_controller.config_port.get.side_effect = lambda key, default=None: (
            "english" if key == "image.ocr_language" else default
        )

        init_cli_locale("zh_CN")
        args = _make_execution_args(action="invoice_cn", to="md", files=[str(src)], ocr=True, ocr_language="japanese")
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.options["ocr_language"] == "japanese"

    def test_to_markdown_request_carries_cli_image_link_style(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "image.png"
        src.write_bytes(b"\x89PNG\r\n\x1a\n")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(to="md", files=[str(src)], image_link_style="markdown_embed")
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.options["image_link_style"] == "markdown_embed"

    def test_to_markdown_request_carries_cli_table_merge_strategy(self, tmp_path) -> None:
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "book.xlsx"
        _write_ooxml(src)

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(to="md", files=[str(src)], table_merge_strategy="empty")
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        assert request.options["table_merge_strategy"] == "empty"

    def test_dry_run_path_uses_same_action(self, tmp_path) -> None:
        """Dry-run path uses the same resolved action."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "test.docx"
        _write_ooxml(src)

        args = _make_execution_args(action="gongwen", dry_run=True, files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            exit_code = execute_convert(args, controller=MagicMock(has_runtime=True))

        assert exit_code == 0

    def test_dry_run_txt_suffix_uses_detected_markdown_workflow(self, tmp_path, capsys) -> None:
        """Dry-run reports detected Markdown while sharing the text workflow."""
        import json
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "notes.txt"
        src.write_text("# Title\n\ncontent\n", encoding="utf-8")

        args = _make_execution_args(json=True, dry_run=True, to="docx", files=[str(src)])
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            exit_code = execute_convert(args, controller=MagicMock(has_runtime=True))

        assert exit_code == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload["data"]["detected_format"] == "markdown"
        assert payload["data"]["routing"] == {
            "status": "deferred_to_runtime",
            "source_format": "markdown",
            "source_category": "markdown",
            "source_candidates": ["markdown"],
            "category_fallback": None,
            "target_format": "docx",
            "action_name": "",
        }
        assert "route_source_format" not in payload["data"]
        assert "route" not in payload["data"]
