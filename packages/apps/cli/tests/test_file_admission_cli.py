"""CLI contract tests for content-first file admission."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from packages.apps.cli.tests.capability_fixtures import bundled_available_runtime_projection

from docwen_cli.commands.execution_v3 import execute_execution
from docwen_cli.main import _build_parser
from docwen_cli.utils import validate_files
from docwen_core.models import (
    FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
    FILE_INSPECTION_METADATA_KEY,
    make_admission_acceptance,
)

pytestmark = pytest.mark.unit


def _available_controller() -> MagicMock:
    controller = MagicMock()
    controller.has_runtime = True
    controller.describe_runtime_capabilities.return_value = bundled_available_runtime_projection()
    return controller


@pytest.mark.parametrize(
    "argv",
    [
        ["convert", "input.md", "--to", "docx", "--output", "out.docx"],
        ["validate", "input.md"],
        ["number", "markdown", "input.md", "--operation", "remove", "--in-place"],
        ["merge", "pdf", "a.pdf", "b.pdf", "--output", "merged.pdf"],
        ["split", "pdf", "input.pdf", "--pages", "1", "--output-dir", "out"],
        ["batch", "validate", "a.md", "b.md"],
    ],
)
def test_every_execution_leaf_accepts_explicit_detected_format(argv: list[str]) -> None:
    args = _build_parser().parse_args([*argv, "--use-detected-format"])

    assert args.use_detected_format is True


def test_cross_family_mismatch_is_rejected_by_default_and_machine_readable(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    source = tmp_path / "layout.docx"
    source.write_bytes(b"%PDF-1.4\n")
    output = tmp_path / "layout.md"
    args = _build_parser().parse_args(
        ["convert", str(source), "--to", "md", "--output", str(output), "--dry-run", "--json"]
    )

    exit_code = execute_execution(args, _available_controller())

    envelope = json.loads(capsys.readouterr().out)
    assert exit_code == 2
    assert envelope["success"] is False
    assert envelope["error"]["category"] == "invalid_input"
    assert envelope["error"]["code"] == "file_format_confirmation_required"
    assert envelope["error"]["details"]["admission"]["relation"] == "cross_family_mismatch"
    assert envelope["error"]["details"]["admission"]["detected_format"] == "pdf"
    assert "--use-detected-format" in envelope["error"]["hint"]


def test_explicit_acceptance_routes_the_detected_content_not_the_suffix(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    source = tmp_path / "layout.docx"
    source.write_bytes(b"%PDF-1.4\n")
    output = tmp_path / "layout.md"
    args = _build_parser().parse_args(
        [
            "convert",
            str(source),
            "--to",
            "md",
            "--output",
            str(output),
            "--dry-run",
            "--json",
            "--use-detected-format",
        ]
    )

    exit_code = execute_execution(args, _available_controller())

    envelope = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert envelope["success"] is True
    assert envelope["data"]["detected_format"] == "pdf"
    assert envelope["data"]["workflow_category"] == "layout"
    assert envelope["data"]["routing"] == {
        "status": "deferred_to_runtime",
        "source_format": "pdf",
        "source_category": "layout",
        "source_candidates": ["pdf", "layout"],
        "category_fallback": "layout",
        "target_format": "md",
        "action_name": "",
    }
    assert envelope["data"]["admission"]["decision"] == "require_explicit_acceptance"
    assert envelope["warnings"][0]["code"] == "FILE_FORMAT_CROSS_FAMILY_MISMATCH"


def test_frozen_runtime_ref_contains_core_acceptance_marker(tmp_path: Path) -> None:
    from docwen_cli.commands.execution_request import file_ref_for_runtime
    from docwen_core.detection import inspect_file

    source = tmp_path / "layout.docx"
    source.write_bytes(b"%PDF-1.4\n")
    inspection = inspect_file(str(source))

    ref = file_ref_for_runtime(
        str(source),
        inspection,
        explicit_acceptance=True,
    )

    assert ref.format == "pdf"
    assert ref.category == "layout"
    assert ref.metadata[FILE_INSPECTION_METADATA_KEY]["detected_format"] == "pdf"
    assert ref.metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY] == make_admission_acceptance(inspection)


@pytest.mark.parametrize("name", ["ordinary.zip", "ordinary.docx"])
def test_generic_zip_remains_blocked_even_with_override(tmp_path: Path, name: str) -> None:
    source = tmp_path / name
    with zipfile.ZipFile(source, "w") as package:
        package.writestr("hello.txt", "hello")

    valid, invalid, warnings = validate_files([str(source)], use_detected_format=True)

    assert valid == []
    assert len(invalid) == 1
    assert "[FILE_CONTAINER_" in invalid[0][1]
    assert warnings == []


def test_unknown_suffix_with_recognizable_content_requires_the_same_explicit_flag(tmp_path: Path) -> None:
    source = tmp_path / "layout.unknown"
    source.write_bytes(b"%PDF-1.4\n")

    default_valid, default_invalid, _ = validate_files([str(source)])
    accepted, invalid, warnings = validate_files([str(source)], use_detected_format=True)

    assert default_valid == []
    assert "[FILE_FORMAT_CONFIRMATION_REQUIRED]" in default_invalid[0][1]
    assert accepted == [str(source.resolve())]
    assert invalid == []
    assert len(warnings) == 1
