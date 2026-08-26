"""Focused tests split from test_cli_contract.py."""

from __future__ import annotations

import pytest

from ._cli_contract_support import (
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

    def test_numbering_options_reach_request(self, tmp_path) -> None:
        """Numbering CLI flags map to converter-contract keys on the request.

        Full user path: ``--clean-numbering remove --add-numbering <scheme>``
        flows through ``build_execution_options`` -> ``normalize_numbering_options``
        and lands on ``ConversionRequest.options`` as the downstream
        converter contract (``remove_numbering`` / ``add_numbering`` /
        ``numbering_scheme``), not as the raw CLI flag names.

        This is the end-to-end mapping assertion the plan's Phase B-CLI
        asks for: typed CLI options -> converter params.
        """
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "test.md"
        src.write_text("# Hello\n", encoding="utf-8")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(
            action="process_md_numbering",
            clean_numbering="remove",
            add_numbering="gongwen_standard",
            files=[str(src)],
        )
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        assert mock_controller.execute_single.called, "controller.execute_single was not invoked"
        request = mock_controller.execute_single.call_args[0][0]
        opts = request.options
        assert request.action_name == "process_md_numbering"

        # Converter-contract keys present with correct values.
        assert opts.get("remove_numbering") is True, (
            f"remove_numbering should be True for --clean-numbering remove, got opts={opts}"
        )
        assert opts.get("add_numbering") is True, (
            f"add_numbering should be True for --add-numbering <scheme>, got opts={opts}"
        )
        assert opts.get("numbering_scheme") == "gongwen_standard", (
            f"numbering_scheme should carry the scheme id verbatim, got opts={opts}"
        )

        # Raw CLI flag names must NOT leak into the converter contract.
        assert "clean_numbering" not in opts, f"raw CLI flag 'clean_numbering' leaked into request.options: {opts}"
        assert "add_numbering" not in opts or opts.get("add_numbering") is True, (
            f"raw CLI flag 'add_numbering' should not appear as a string: {opts}"
        )
        assert "locale" not in opts
        assert "yaml_key_labels" not in opts

    def test_numbering_keep_only_no_removal(self, tmp_path) -> None:
        """``--clean-numbering keep`` maps to remove_numbering=False on the request."""
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "test.md"
        src.write_text("# Hello\n", encoding="utf-8")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(
            action="process_md_numbering",
            clean_numbering="keep",
            files=[str(src)],
        )
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        opts = request.options
        assert opts.get("remove_numbering") is False, (
            f"--clean-numbering keep must map to remove_numbering=False, got opts={opts}"
        )
        assert opts.get("add_numbering") is False, (
            f"no --add-numbering must map to add_numbering=False, got opts={opts}"
        )

    def test_render_mode_reaches_request(self, tmp_path) -> None:
        """--heading-numbering-render-mode word_native lands on the request.

        Full CLI user path for the new Phase B-CLI-render param:
        ``--heading-numbering-render-mode word_native`` flows through
        ``build_execution_options`` -> ``normalize_numbering_options``
        and lands on ``ConversionRequest.options``.
        """
        from unittest.mock import patch

        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "test.md"
        src.write_text("# Hello\n", encoding="utf-8")

        mock_controller = MagicMock()
        mock_controller.has_runtime = True
        mock_controller.execute_single.return_value = MagicMock(success=True)

        args = _make_execution_args(
            to="docx",
            add_numbering="gongwen_standard",
            heading_numbering_render_mode="word_native",
            files=[str(src)],
        )
        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(src)], [], []),
        ):
            execute_convert(args, controller=mock_controller)

        request = mock_controller.execute_single.call_args[0][0]
        opts = request.options
        assert opts.get("heading_numbering_render_mode") == "word_native", (
            f"--heading-numbering-render-mode word_native must reach request.options, got {opts}"
        )


class TestJsonOutputCommandActionSeparation:
    """JSON output separates envelope ``command`` from data ``action_name``."""

    def test_single_presenter_separates_command_and_action(self, capsys) -> None:
        """present_single puts command='convert' in envelope, action_name in data."""
        from docwen_cli.presenters.json_presenter import JsonPresenter
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionMetrics, ConversionResult

        result = ConversionResult(
            task_id="t1",
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id="art-1",
                    kind="primary",
                    staging_path="/tmp/out.md",
                    suggested_name="out.md",
                    media_type="text/markdown",
                    is_primary=True,
                )
            ],
            metrics=ConversionMetrics(duration_ms=100.0),
        )

        presenter = JsonPresenter()
        presenter.present_single(result, command="convert", action_name="gongwen")

        import json

        captured = capsys.readouterr()
        data = json.loads(captured.out)

        assert data["command"] == "convert"
        assert "action_name" not in data["data"]

    def test_single_with_empty_action(self, capsys) -> None:
        """present_single with empty action_name."""
        from docwen_cli.presenters.json_presenter import JsonPresenter
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionMetrics, ConversionResult

        result = ConversionResult(
            task_id="t1",
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id="art-1",
                    kind="primary",
                    staging_path="/tmp/out.md",
                    suggested_name="out.md",
                    media_type="text/markdown",
                    is_primary=True,
                )
            ],
            metrics=ConversionMetrics(duration_ms=100.0),
        )

        presenter = JsonPresenter()
        presenter.present_single(result, command="convert", action_name="")

        import json

        captured = capsys.readouterr()
        data = json.loads(captured.out)

        assert data["command"] == "convert"
        assert "action_name" not in data["data"]

    def test_batch_separates_command_and_action(self, capsys) -> None:
        """present_batch puts command in envelope, action_name in data."""
        from docwen_cli.presenters.json_presenter import JsonPresenter
        from docwen_core.models.result import ConversionMetrics, ConversionResult

        result = ConversionResult(
            task_id="t1",
            success=True,
            metrics=ConversionMetrics(duration_ms=100.0),
        )

        presenter = JsonPresenter()
        presenter.present_batch([result], command="convert", action_name="merge_pdfs")

        import json

        captured = capsys.readouterr()
        data = json.loads(captured.out)

        assert data["command"] == "convert"
        assert "action_name" not in data["data"]

    def test_error_presenter_uses_command(self, capsys) -> None:
        """present_error uses the CLI command name."""
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter()
        presenter.present_error("convert", "test error", error_code="invalid_input")

        import json

        captured = capsys.readouterr()
        data = json.loads(captured.out)

        assert data["command"] == "convert"

    def test_convert_invalid_input_error_uses_convert_command(self, capsys) -> None:
        """convert --json errors keep the CLI command in the envelope command field."""
        from docwen_cli.commands.convert import _print_invalid_input

        args = _make_execution_args(json=True, action="merge_pdfs")

        exit_code = _print_invalid_input("merge_pdfs", args, "bad input")

        import json

        captured = capsys.readouterr()
        data = json.loads(captured.out)

        assert exit_code != 0
        assert data["command"] == "convert"
        assert data["error"]["code"] == "invalid_input"

    def test_convert_runtime_unavailable_error_uses_convert_command(self, capsys) -> None:
        from docwen_cli.commands.convert import _print_unavailable

        args = _make_execution_args(json=True, action="gongwen")

        exit_code = _print_unavailable("gongwen", args, "runtime unavailable")

        import json

        captured = capsys.readouterr()
        data = json.loads(captured.out)

        assert exit_code != 0
        assert data["command"] == "convert"
        assert data["error"]["code"] == "dependency_missing"

    def test_dry_run_json_uses_convert_command(self, tmp_path, capsys) -> None:
        from docwen_cli.commands.convert import execute_convert

        src = tmp_path / "note.docx"
        _write_ooxml(src)
        args = _make_execution_args(json=True, dry_run=True, action="gongwen", files=[str(src)], to="md")

        exit_code = execute_convert(args, controller=MagicMock(has_runtime=True))

        import json

        captured = capsys.readouterr()
        data = json.loads(captured.out)

        assert exit_code == 0
        assert data["command"] == "convert"
        assert data["data"]["routing"]["source_format"] == "docx"
        assert data["data"]["routing"]["status"] == "deferred_to_runtime"
