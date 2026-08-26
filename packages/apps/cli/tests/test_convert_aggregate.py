"""Unit tests for CLI aggregate (merge) execution path.

Tests verify:
- _execute_aggregate builds a single ConversionRequest with all input_refs
- _execute_aggregate fails for fewer than 2 valid files
- execute_convert routes merge-* actions to aggregate path
- is_aggregate_action integration in CLI context
"""

from __future__ import annotations

import os
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from docwen_application.runtime_capability_catalog import RuntimeRoute
from docwen_core.models import (
    AdmissionDecision,
    DetectionConfidence,
    DetectionMethod,
    FileInspection,
    FormatRelation,
    StructureStatus,
)

pytestmark = pytest.mark.unit


# ── Helpers ─────────────────────────────────────────────────────────────────


def _fake_args(extra: dict | None = None) -> object:
    """Build a minimal namespace for normalized aggregate commands."""
    import argparse

    ns = argparse.Namespace()
    # Common args
    ns.action = ""
    ns.json = False
    ns.quiet = False
    ns.verbose = False
    ns.timing = False
    ns.batch = False
    ns.jobs = 1
    ns.continue_on_error = False
    ns.output = None
    ns.dry_run = False
    ns.template = None
    ns.check = []
    ns.extract_img = False
    ns.no_extract_img = False
    ns.ocr = False
    ns.image_mode = None
    ns.ocr_placement = None
    ns.clean_numbering = None
    ns.add_numbering = None
    ns.heading_merge_mode = None
    # Files
    ns.files = ["/test/a.pdf", "/test/b.pdf"]
    # Aggregate-specific
    ns.pages = None
    ns.dpi = None
    ns.mode = None
    ns.keep_alpha = False

    if extra:
        for k, v in extra.items():
            setattr(ns, k, v)
    return ns


_FORMAT_BY_SUFFIX = {
    ".csv": ("csv", "spreadsheet"),
    ".jpg": ("jpeg", "image"),
    ".pdf": ("pdf", "layout"),
    ".png": ("png", "image"),
    ".xlsx": ("xlsx", "spreadsheet"),
}


def _inspection(path: str) -> FileInspection:
    """Build a frozen valid inspection for request-construction unit tests."""

    absolute = os.path.abspath(path)
    extension = Path(path).suffix.lower()
    detected_format, category = _FORMAT_BY_SUFFIX[extension]
    declared_format = "jpeg" if extension == ".jpg" else extension.lstrip(".")
    relation = FormatRelation.EQUIVALENT_ALIAS if extension == ".jpg" else FormatRelation.EXACT_MATCH
    return FileInspection(
        file_path=absolute,
        size_bytes=16,
        mtime_ns=0,
        extension=extension,
        declared_format=declared_format,
        declared_category=category,
        detected_format=detected_format,
        detected_category=category,
        workflow_category=category,
        detection_method=DetectionMethod.SIGNATURE,
        confidence=DetectionConfidence.CERTAIN,
        structure_status=StructureStatus.VALID,
        relation=relation,
        decision=AdmissionDecision.ALLOW,
        declared_supported=True,
        detected_supported=True,
    )


def _inspection_map(files: list[str]) -> dict[str, FileInspection]:
    return {os.path.abspath(path): _inspection(path) for path in files}


def _aggregate_routes(files: list[str], action: str, target: str) -> dict[str, RuntimeRoute]:
    resolved_target = "tif" if action == "merge_images_to_tiff" else target
    route_id = f"test:{action}:{resolved_target}"
    return {
        file_path: RuntimeRoute(
            id=route_id,
            operation="action",
            source=_inspection(file_path).workflow_category,
            source_category=_inspection(file_path).workflow_category,
            target=resolved_target,
            action_name=action,
            available=True,
            state="available",
            options=("merge_mode",) if action == "merge_tables" else (),
        )
        for file_path in files
    }


def _accept_files(files: list[str], **kwargs: object) -> tuple[list[str], list[object], list[object]]:
    """Stand in for validation while preserving its frozen-inspection output."""

    cache = kwargs.get("inspection_cache")
    assert isinstance(cache, dict)
    cache.update(_inspection_map(files))
    return files, [], []


def _aggregate_capability_projection() -> dict[str, object]:
    """Return the canonical category routes used by aggregate integration tests."""

    routes_by_source: list[tuple[str, str, dict[str, object]]] = [
        (
            "pdf",
            "layout",
            {
                "id": "print:pdf:pdf:merge_pdfs",
                "operation": "action",
                "source": "pdf",
                "target": "pdf",
                "action": "merge_pdfs",
                "available": True,
                "state": "available",
                "options": [],
            },
        ),
        (
            "spreadsheet",
            "spreadsheet",
            {
                "id": "spreadsheet:spreadsheet:xlsx:merge_tables",
                "operation": "action",
                "source": "spreadsheet",
                "target": "xlsx",
                "action": "merge_tables",
                "available": True,
                "state": "available",
                "options": ["merge_mode"],
            },
        ),
        (
            "image",
            "image",
            {
                "id": "image:image:tif:merge_images_to_tiff",
                "operation": "action",
                "source": "image",
                "target": "tif",
                "action": "merge_images_to_tiff",
                "available": True,
                "state": "available",
                "options": [],
            },
        ),
    ]
    sources = [
        {"id": source_id, "category": category, "available": True, "routes": [route]}
        for source_id, category, route in routes_by_source
    ]
    return {
        "resource": "formats",
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
        "security": {"dependency_egress_guard": {}},
        "gates": [],
        "sources": sources,
        "counts": {
            "sources": 3,
            "routes": 3,
            "available_routes": 3,
            "unavailable_routes": 0,
            "actions": 3,
        },
    }


def _aggregate_controller() -> MagicMock:
    controller = MagicMock()
    controller.has_runtime = True
    controller.describe_runtime_capabilities.return_value = _aggregate_capability_projection()
    controller.execute_aggregate.return_value = MagicMock(success=True)
    return controller


# ── is_aggregate_action integration ─────────────────────────────────────────


class TestIsAggregateActionCLI:
    """Verify is_aggregate_action is importable and correct in CLI context."""

    def test_import_from_application(self) -> None:
        from docwen_application.commands.batch import is_aggregate_action

        assert callable(is_aggregate_action)

    def test_merge_pdfs_is_aggregate(self) -> None:
        from docwen_application.commands.batch import is_aggregate_action

        assert is_aggregate_action("merge_pdfs") is True

    def test_convert_is_not_aggregate(self) -> None:
        from docwen_application.commands.batch import is_aggregate_action

        assert is_aggregate_action("convert") is False


# ── _execute_aggregate helper ───────────────────────────────────────────────


class TestExecuteAggregate:
    """Test the _execute_aggregate() helper in convert.py."""

    def test_builds_multi_input_request(self) -> None:
        """_execute_aggregate builds a single request with all input_refs."""

        from docwen_cli.commands.convert import _execute_aggregate
        from docwen_core.models.result import ConversionResult

        mock_controller = MagicMock()
        mock_controller.execute_aggregate.return_value = ConversionResult(task_id="agg-1", success=True)

        args = _fake_args(
            {
                "command": "merge-pdfs",
                "files": ["/test/a.pdf", "/test/b.pdf", "/test/c.pdf"],
                "output": "/out/merged.pdf",
            }
        )

        exit_code = _execute_aggregate(
            mock_controller,
            action="merge_pdfs",
            files=["/test/a.pdf", "/test/b.pdf", "/test/c.pdf"],
            target_format="pdf",
            options={},
            args=args,  # type: ignore[arg-type]
            inspections=_inspection_map(["/test/a.pdf", "/test/b.pdf", "/test/c.pdf"]),
            routes_by_file=_aggregate_routes(["/test/a.pdf", "/test/b.pdf", "/test/c.pdf"], "merge_pdfs", "pdf"),
        )

        mock_controller.execute_aggregate.assert_called_once()
        called_request, called_action = mock_controller.execute_aggregate.call_args.args
        assert called_action == "merge_pdfs"
        assert len(called_request.input_refs) == 3
        assert exit_code == 0

    def test_merge_tables_request_preserves_concrete_formats_and_spreadsheet_category(self) -> None:
        """Aggregate refs retain content formats plus their category fallback."""

        from docwen_cli.commands.convert import _execute_aggregate
        from docwen_core.models.result import ConversionResult

        mock_controller = MagicMock()
        mock_controller.execute_aggregate.return_value = ConversionResult(task_id="agg-tables", success=True)

        args = _fake_args(
            {
                "command": "merge",
                "command_path": "merge tables",
                "action": "merge_tables",
                "files": ["/test/a.xlsx", "/test/b.csv"],
                "mode": "cell",
            }
        )

        exit_code = _execute_aggregate(
            mock_controller,
            action="merge_tables",
            files=["/test/a.xlsx", "/test/b.csv"],
            target_format="xlsx",
            options={"merge_mode": "cell"},
            args=args,  # type: ignore[arg-type]
            inspections=_inspection_map(["/test/a.xlsx", "/test/b.csv"]),
            routes_by_file=_aggregate_routes(["/test/a.xlsx", "/test/b.csv"], "merge_tables", "xlsx"),
        )

        called_request, called_action = mock_controller.execute_aggregate.call_args.args
        assert exit_code == 0
        assert called_action == "merge_tables"
        assert called_request.action_name == "merge_tables"
        assert called_request.target_format == "xlsx"
        assert [ref.format for ref in called_request.input_refs] == ["xlsx", "csv"]
        assert [ref.category for ref in called_request.input_refs] == ["spreadsheet", "spreadsheet"]

    def test_merge_images_to_tiff_request_preserves_formats_and_image_category(self) -> None:
        """Public TIFF wording keeps concrete inputs and the image fallback."""

        from docwen_cli.commands.convert import _execute_aggregate
        from docwen_core.models.result import ConversionResult

        mock_controller = MagicMock()
        mock_controller.execute_aggregate.return_value = ConversionResult(task_id="agg-images", success=True)

        args = _fake_args(
            {
                "command": "merge",
                "command_path": "merge images",
                "action": "merge_images_to_tiff",
                "files": ["/test/a.png", "/test/b.jpg"],
            }
        )

        exit_code = _execute_aggregate(
            mock_controller,
            action="merge_images_to_tiff",
            files=["/test/a.png", "/test/b.jpg"],
            target_format="tiff",
            options={},
            args=args,  # type: ignore[arg-type]
            inspections=_inspection_map(["/test/a.png", "/test/b.jpg"]),
            routes_by_file=_aggregate_routes(["/test/a.png", "/test/b.jpg"], "merge_images_to_tiff", "tiff"),
        )

        called_request, called_action = mock_controller.execute_aggregate.call_args.args
        assert exit_code == 0
        assert called_action == "merge_images_to_tiff"
        assert called_request.action_name == "merge_images_to_tiff"
        assert called_request.target_format == "tif"
        assert [ref.format for ref in called_request.input_refs] == ["png", "jpeg"]
        assert [ref.category for ref in called_request.input_refs] == ["image", "image"]

    def test_fewer_than_two_files_returns_localized_error(
        self,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """_execute_aggregate rejects single-file aggregate requests."""
        from docwen_cli.commands.convert import _execute_aggregate

        mock_controller = MagicMock()

        args = _fake_args(
            {
                "command": "merge-pdfs",
                "files": ["/test/only.pdf"],
            }
        )

        exit_code = _execute_aggregate(
            mock_controller,
            action="merge_pdfs",
            files=["/test/only.pdf"],
            target_format="pdf",
            options={},
            args=args,  # type: ignore[arg-type]
            routes_by_file=_aggregate_routes(["/test/only.pdf"], "merge_pdfs", "pdf"),
        )

        # Should return non-zero exit code
        assert exit_code != 0
        # Should NOT try to execute
        mock_controller.execute_aggregate.assert_not_called()
        stderr = capsys.readouterr().err
        assert "cli.messages.error_aggregate_need_two" not in stderr
        assert "{count}" not in stderr
        assert "1" in stderr

    def test_failure_result_returns_error_code(self) -> None:
        """_execute_aggregate returns error code when conversion fails."""
        from docwen_cli.commands.convert import _execute_aggregate
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        mock_controller = MagicMock()
        mock_controller.execute_aggregate.return_value = ConversionResult(
            task_id="agg-fail",
            success=False,
            error=ConversionErrorInfo(error_type="conversion_failed", message="merge error"),
        )

        args = _fake_args(
            {
                "command": "merge-pdfs",
                "files": ["/test/a.pdf", "/test/b.pdf"],
            }
        )

        exit_code = _execute_aggregate(
            mock_controller,
            action="merge_pdfs",
            files=["/test/a.pdf", "/test/b.pdf"],
            target_format="pdf",
            options={},
            args=args,  # type: ignore[arg-type]
            inspections=_inspection_map(["/test/a.pdf", "/test/b.pdf"]),
            routes_by_file=_aggregate_routes(["/test/a.pdf", "/test/b.pdf"], "merge_pdfs", "pdf"),
        )

        assert exit_code != 0

    def test_exception_returns_error_code(self) -> None:
        """_execute_aggregate handles exceptions gracefully."""
        from docwen_cli.commands.convert import _execute_aggregate

        mock_controller = MagicMock()
        mock_controller.execute_aggregate.side_effect = RuntimeError("runtime down")

        args = _fake_args(
            {
                "command": "merge-pdfs",
                "files": ["/test/a.pdf", "/test/b.pdf"],
            }
        )

        exit_code = _execute_aggregate(
            mock_controller,
            action="merge_pdfs",
            files=["/test/a.pdf", "/test/b.pdf"],
            target_format="pdf",
            options={},
            args=args,  # type: ignore[arg-type]
            inspections=_inspection_map(["/test/a.pdf", "/test/b.pdf"]),
            routes_by_file=_aggregate_routes(["/test/a.pdf", "/test/b.pdf"], "merge_pdfs", "pdf"),
        )

        assert exit_code != 0

    def test_json_mode_presenter_called(self) -> None:
        """_execute_aggregate uses JsonPresenter when --json is set."""
        from docwen_cli.commands.convert import _execute_aggregate
        from docwen_core.models.result import ConversionResult

        mock_controller = MagicMock()
        mock_controller.execute_aggregate.return_value = ConversionResult(task_id="agg-json", success=True)

        args = _fake_args(
            {
                "command": "merge-pdfs",
                "files": ["/test/a.pdf", "/test/b.pdf"],
                "json": True,
            }
        )

        exit_code = _execute_aggregate(
            mock_controller,
            action="merge_pdfs",
            files=["/test/a.pdf", "/test/b.pdf"],
            target_format="pdf",
            options={},
            args=args,  # type: ignore[arg-type]
            json_mode=True,
            inspections=_inspection_map(["/test/a.pdf", "/test/b.pdf"]),
            routes_by_file=_aggregate_routes(["/test/a.pdf", "/test/b.pdf"], "merge_pdfs", "pdf"),
        )

        assert exit_code == 0

    def test_input_refs_preserve_format_and_category(self) -> None:
        """Each FileRef in the aggregate request includes format and category."""
        from docwen_cli.commands.convert import _execute_aggregate
        from docwen_core.models.result import ConversionResult

        mock_controller = MagicMock()
        mock_controller.execute_aggregate.return_value = ConversionResult(task_id="agg-ref", success=True)

        args = _fake_args(
            {
                "command": "merge-tables",
                "files": ["/test/t1.xlsx", "/test/t2.xlsx"],
            }
        )

        _execute_aggregate(
            mock_controller,
            action="merge_tables",
            files=["/test/t1.xlsx", "/test/t2.xlsx"],
            target_format="xlsx",
            options={},
            args=args,  # type: ignore[arg-type]
            inspections=_inspection_map(["/test/t1.xlsx", "/test/t2.xlsx"]),
            routes_by_file=_aggregate_routes(["/test/t1.xlsx", "/test/t2.xlsx"], "merge_tables", "xlsx"),
        )

        called_request = mock_controller.execute_aggregate.call_args.args[0]
        assert len(called_request.input_refs) == 2
        for ref in called_request.input_refs:
            assert ref.format  # format should be populated
            assert ref.category  # category should be populated


# ── Aggregate action routing in execute_convert ─────────────────────────────


class TestExecuteConvertAggregateRouting:
    """Verify execute_convert() routes aggregate actions correctly."""

    def test_merge_pdfs_routes_to_aggregate(self) -> None:
        """The public PDF merge path uses the aggregate controller boundary."""

        from docwen_cli.commands.convert import execute_convert

        mock_controller = _aggregate_controller()

        args = _fake_args(
            {
                "command": "merge",
                "command_path": "merge pdf",
                "action": "merge_pdfs",
                "files": ["/test/a.pdf", "/test/b.pdf"],
            }
        )

        # Mock validate_files to return all files as valid (test paths don't exist)
        with patch(
            "docwen_cli.commands.convert.validate_files",
            side_effect=_accept_files,
        ):
            exit_code = execute_convert(args, controller=mock_controller)  # type: ignore[arg-type]

        assert exit_code == 0
        assert mock_controller.execute_aggregate.call_args.args[1] == "merge_pdfs"

    def test_merge_tables_routes_to_aggregate(self) -> None:
        """The public table merge path uses the aggregate controller boundary."""

        from docwen_cli.commands.convert import execute_convert

        mock_controller = _aggregate_controller()

        args = _fake_args(
            {
                "command": "merge",
                "command_path": "merge tables",
                "action": "merge_tables",
                "files": ["/test/t1.xlsx", "/test/t2.xlsx", "/test/t3.xlsx"],
            }
        )

        with patch(
            "docwen_cli.commands.convert.validate_files",
            side_effect=_accept_files,
        ):
            exit_code = execute_convert(args, controller=mock_controller)  # type: ignore[arg-type]

        assert exit_code == 0
        assert mock_controller.execute_aggregate.call_args.args[1] == "merge_tables"

    def test_merge_images_to_tiff_routes_to_aggregate(self) -> None:
        """merge_images_to_tiff with >=2 files routes to aggregate."""

        from docwen_cli.commands.convert import execute_convert

        mock_controller = _aggregate_controller()

        args = _fake_args(
            {
                "command": "merge",
                "command_path": "merge images",
                "action": "merge_images_to_tiff",
                "files": ["/test/a.png", "/test/b.png"],
            }
        )

        with patch(
            "docwen_cli.commands.convert.validate_files",
            side_effect=_accept_files,
        ):
            exit_code = execute_convert(args, controller=mock_controller)  # type: ignore[arg-type]

        assert exit_code == 0
        assert mock_controller.execute_aggregate.call_args.args[1] == "merge_images_to_tiff"

    def test_aggregate_with_insufficient_files_rejected(self) -> None:
        """aggregate action with only 1 valid file returns error."""

        from docwen_cli.commands.convert import execute_convert

        mock_controller = MagicMock()
        mock_controller.has_runtime = True

        args = _fake_args(
            {
                "command": "merge",
                "command_path": "merge pdf",
                "action": "merge_pdfs",
                "files": ["/test/only.pdf"],
            }
        )

        with patch(
            "docwen_cli.commands.convert.validate_files",
            side_effect=_accept_files,
        ):
            exit_code = execute_convert(args, controller=mock_controller)  # type: ignore[arg-type]

        # Should get an error exit code (need at least 2 files for aggregate)
        assert exit_code != 0
        mock_controller.execute_aggregate.assert_not_called()

    def test_aggregate_no_runtime_returns_error(self) -> None:
        """merge-pdfs without a runtime returns error."""

        from docwen_cli.commands.convert import execute_convert

        mock_controller = MagicMock()
        mock_controller.has_runtime = False

        args = _fake_args(
            {
                "action": "merge_pdfs",
                "files": ["/test/a.pdf", "/test/b.pdf"],
            }
        )

        with patch(
            "docwen_cli.commands.convert.validate_files",
            side_effect=_accept_files,
        ):
            exit_code = execute_convert(args, controller=mock_controller)  # type: ignore[arg-type]

        assert exit_code != 0
