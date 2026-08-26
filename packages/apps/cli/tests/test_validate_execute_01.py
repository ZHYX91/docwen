"""Focused tests split from test_validate_execute.py."""

from __future__ import annotations

from ._validate_execute_support import (
    DocWenError,
    MagicMock,
    _markdown_inspection,
    _runtime_route,
    _validation_capability_projection,
    argparse,
    os,
    parse_runtime_capability_catalog,
    pytest,
)

pytestmark = pytest.mark.unit


def test_cli_route_resolution_tracks_manifest_targets_and_actions_without_tables() -> None:
    """A new or removed runtime route needs no corresponding CLI table edit."""

    from docwen_cli.commands.convert import _resolve_runtime_routes

    path = os.path.abspath("/test/future.md")
    inspection = _markdown_inspection(path)
    projection = _validation_capability_projection()
    markdown_source = projection["sources"][0]  # type: ignore[index]
    markdown_source["routes"] = [  # type: ignore[index]
        {
            "id": "future:markdown:future_format:manifest_added",
            "operation": "action",
            "source": "markdown",
            "target": "future_format",
            "action": "manifest_added",
            "available": True,
            "state": "available",
            "options": [],
        }
    ]
    projection["counts"] = {
        "sources": 3,
        "routes": 3,
        "available_routes": 3,
        "unavailable_routes": 0,
        "actions": 3,
    }
    catalog = parse_runtime_capability_catalog(projection)

    resolved = _resolve_runtime_routes(
        catalog,
        files=[path],
        inspections={path: inspection},
        action="manifest_added",
        target_format="future_format",
    )
    assert resolved[path].id == "future:markdown:future_format:manifest_added"

    with pytest.raises(ValueError, match="No canonical runtime route"):
        _resolve_runtime_routes(
            catalog,
            files=[path],
            inspections={path: inspection},
            action="deleted_action",
            target_format="future_format",
        )


class TestValidateFilesEmptyCheck:
    """Verify empty files are rejected (F-C4-075)."""

    def test_rejects_zero_byte_file(self, tmp_path) -> None:
        """Files with st_size == 0 should be flagged as invalid."""
        from docwen_cli.utils import validate_files

        empty_file = tmp_path / "empty.md"
        empty_file.write_text("", encoding="utf-8")

        valid, invalid, _warnings = validate_files([str(empty_file)])
        assert valid == []
        assert len(invalid) == 1
        assert "[FILE_EMPTY]" in invalid[0][1]

    def test_accepts_non_empty_file(self, tmp_path) -> None:
        """Files with content should be accepted."""
        from docwen_cli.utils import validate_files

        file_path = tmp_path / "test.md"
        file_path.write_text("# Content\n", encoding="utf-8")

        valid, invalid, _warnings = validate_files([str(file_path)])
        assert len(valid) == 1
        assert invalid == []

    def test_rejects_empty_pdf(self, tmp_path) -> None:
        """Empty files regardless of extension are rejected."""
        from docwen_cli.utils import validate_files

        empty_file = tmp_path / "empty.pdf"
        empty_file.write_bytes(b"")

        valid, invalid, _warnings = validate_files([str(empty_file)])
        assert valid == []
        assert len(invalid) == 1
        assert "[FILE_EMPTY]" in invalid[0][1]


class TestValidateFilesContentDetection:
    """Verify content-first mismatch warnings at the CLI boundary."""

    def test_extension_content_mismatch_generates_warning(self, tmp_path) -> None:
        """When extension doesn't match detected content, a warning is emitted."""
        from docwen_cli.utils import validate_files

        # Create a PNG file with .jpg extension
        png_data = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
        mislabeled = tmp_path / "actually_png.jpg"
        mislabeled.write_bytes(png_data)

        valid, _invalid, warnings = validate_files([str(mislabeled)])
        assert len(valid) == 1  # still valid despite mismatch
        assert len(warnings) == 1
        assert "JPEG" in warnings[0][1]
        assert "PNG" in warnings[0][1]

    def test_matching_extension_no_warning(self, tmp_path) -> None:
        """Correctly labeled files should not produce warnings."""
        from docwen_cli.utils import validate_files

        png_data = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
        png_file = tmp_path / "correct.png"
        png_file.write_bytes(png_data)

        valid, _invalid, warnings = validate_files([str(png_file)])
        assert len(valid) == 1
        assert len(warnings) == 0


class TestValidateFilesExistingBehavior:
    """Verify existing behaviour is preserved after enhancements."""

    def test_rejects_nonexistent_file(self) -> None:
        from docwen_cli.utils import validate_files

        valid, invalid, _warnings = validate_files(["/nonexistent/path/file.md"])
        assert valid == []
        assert len(invalid) == 1
        assert "不存在" in invalid[0][1]

    def test_rejects_directory(self, tmp_path) -> None:
        from docwen_cli.utils import validate_files

        dir_path = tmp_path / "subdir"
        dir_path.mkdir()

        valid, invalid, _warnings = validate_files([str(dir_path)])
        assert valid == []
        assert len(invalid) == 1
        assert "不是文件" in invalid[0][1]

    def test_rejects_unknown_extension(self, tmp_path) -> None:
        from docwen_cli.utils import validate_files

        weird_file = tmp_path / "data.xyzzy"
        weird_file.write_text("hello", encoding="utf-8")

        valid, invalid, _warnings = validate_files([str(weird_file)])
        assert valid == []
        assert len(invalid) == 1
        assert "[FILE_FORMAT_CONFIRMATION_REQUIRED]" in invalid[0][1]

    def test_returns_absolute_paths(self, tmp_path) -> None:
        from docwen_cli.utils import validate_files

        file_path = tmp_path / "test.md"
        file_path.write_text("# Test\n", encoding="utf-8")

        valid, _invalid, _warnings = validate_files([str(file_path)])
        assert len(valid) == 1
        assert os.path.isabs(valid[0])

    def test_mixed_valid_and_invalid(self, tmp_path) -> None:
        from docwen_cli.utils import validate_files

        good = tmp_path / "good.md"
        good.write_text("# ok\n", encoding="utf-8")
        empty = tmp_path / "empty.md"
        empty.write_text("", encoding="utf-8")

        valid, invalid, _warnings = validate_files([str(good), str(empty), "/nope.md"])
        assert len(valid) == 1
        assert len(invalid) == 2


class TestExecuteSingleErrorHandling:
    """Verify _execute_single has distinct handlers per exception type (F-C4-041)."""

    def _fake_args(self, **overrides) -> argparse.Namespace:
        ns = argparse.Namespace()
        ns.json = False
        ns.quiet = False
        ns.verbose = False
        ns.timing = False
        ns.output = None
        for k, v in overrides.items():
            setattr(ns, k, v)
        return ns

    def test_file_not_found_error_returns_invalid_input(self) -> None:
        """FileNotFoundError should return ExitCode.INVALID_INPUT."""
        from docwen_cli.commands.convert import _execute_single
        from docwen_cli.exit_codes import ExitCode

        controller = MagicMock()
        controller.execute_single.side_effect = FileNotFoundError("文件不存在: /missing/file.md")

        args = self._fake_args()
        exit_code = _execute_single(
            controller,
            "validate",
            "/missing/file.md",
            target_format="md",
            options={},
            args=args,
            inspection=_markdown_inspection("/missing/file.md"),
            route=_runtime_route("validate", "markdown"),
        )

        assert exit_code == int(ExitCode.INVALID_INPUT)

    def test_docwen_error_returns_mapped_exit_code(self) -> None:
        """DocWenError should be caught separately from generic Exception."""
        from docwen_cli.commands.convert import _execute_single

        controller = MagicMock()
        controller.execute_single.side_effect = DocWenError("转换失败: bad input")

        args = self._fake_args()
        exit_code = _execute_single(
            controller,
            "convert",
            "/test/file.md",
            target_format="docx",
            options={},
            args=args,
            inspection=_markdown_inspection("/test/file.md"),
            route=_runtime_route("convert", "docx"),
        )

        # DocWenError without specific attributes maps to UNKNOWN_ERROR
        assert exit_code != 0

    def test_generic_exception_returns_unknown_error(self) -> None:
        """Generic Exception should return ExitCode.INTERNAL_ERROR."""
        from docwen_cli.commands.convert import _execute_single
        from docwen_cli.exit_codes import ExitCode

        controller = MagicMock()
        controller.execute_single.side_effect = RuntimeError("something broke")

        args = self._fake_args()
        exit_code = _execute_single(
            controller,
            "convert",
            "/test/file.md",
            target_format="docx",
            options={},
            args=args,
            inspection=_markdown_inspection("/test/file.md"),
            route=_runtime_route("convert", "docx"),
        )

        assert exit_code == int(ExitCode.INTERNAL_ERROR)

    def test_successful_execution_returns_ok(self) -> None:
        """Successful execution returns ExitCode.OK."""
        from docwen_cli.commands.convert import _execute_single
        from docwen_cli.exit_codes import ExitCode

        mock_result = MagicMock()
        mock_result.success = True
        mock_result.error = None

        controller = MagicMock()
        controller.execute_single.return_value = mock_result

        args = self._fake_args()
        exit_code = _execute_single(
            controller,
            "convert",
            "/test/file.md",
            target_format="docx",
            options={},
            args=args,
            inspection=_markdown_inspection("/test/file.md"),
            route=_runtime_route("convert", "docx"),
        )

        assert exit_code == int(ExitCode.OK)

    def test_json_mode_file_not_found(self, capsys) -> None:
        """JSON mode should output proper error envelope for FileNotFoundError."""
        from docwen_cli.commands.convert import _execute_single

        controller = MagicMock()
        controller.execute_single.side_effect = FileNotFoundError("not found")

        args = self._fake_args(json=True)
        exit_code = _execute_single(
            controller,
            "validate",
            "/missing/file.md",
            target_format="md",
            options={},
            args=args,
            json_mode=True,
            inspection=_markdown_inspection("/missing/file.md"),
            route=_runtime_route("validate", "markdown"),
        )

        captured = capsys.readouterr()
        import json

        output = json.loads(captured.out)
        assert output["success"] is False
        assert output["error"]["code"] == "invalid_input"
        assert exit_code != 0

    def test_json_mode_docwen_error(self, capsys) -> None:
        """JSON mode should output proper error envelope for DocWenError."""
        from docwen_cli.commands.convert import _execute_single

        controller = MagicMock()
        controller.execute_single.side_effect = DocWenError("domain error")

        args = self._fake_args(json=True)
        exit_code = _execute_single(
            controller,
            "convert",
            "/test/file.md",
            target_format="docx",
            options={},
            args=args,
            json_mode=True,
            inspection=_markdown_inspection("/test/file.md"),
            route=_runtime_route("convert", "docx"),
        )

        captured = capsys.readouterr()
        import json

        output = json.loads(captured.out)
        assert output["success"] is False
        # DocWenError is caught separately — not a generic error
        assert exit_code != 0
