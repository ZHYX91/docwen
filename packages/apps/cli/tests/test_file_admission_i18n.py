"""CLI localization tests for stable file-admission diagnostics."""

from __future__ import annotations

import json
from collections.abc import Iterator
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from packages.apps.cli.tests.capability_fixtures import bundled_available_runtime_projection

from docwen_cli.commands.execution_v3 import execute_execution
from docwen_cli.commands.inspect import execute_inspect
from docwen_cli.file_admission_i18n import render_file_admission_code
from docwen_cli.i18n import get_cli_locale, init_cli_locale
from docwen_cli.main import _build_parser
from docwen_cli.utils import validate_files

pytestmark = pytest.mark.unit


def _available_controller() -> MagicMock:
    controller = MagicMock()
    controller.has_runtime = True
    controller.describe_runtime_capabilities.return_value = bundled_available_runtime_projection()
    return controller


@pytest.fixture(autouse=True)
def _restore_cli_locale() -> Iterator[None]:
    previous = get_cli_locale()
    try:
        yield
    finally:
        init_cli_locale(previous)


def test_get_cli_locale_tracks_supported_setting_and_ignores_unknown_value() -> None:
    """The public getter exposes only a successfully selected locale."""
    init_cli_locale("de_DE")
    assert get_cli_locale() == "de_DE"

    init_cli_locale("not-a-supported-locale")
    assert get_cli_locale() == "de_DE"


@pytest.mark.parametrize(
    ("locale", "expected"),
    [
        ("zh_CN", "所选文件为空。"),
        ("de_DE", "Die ausgewählte Datei ist leer."),
        ("ja_JP", "選択したファイルは空です。"),
    ],
)
def test_adapter_uses_runtime_file_admission_catalog(locale: str, expected: str) -> None:
    init_cli_locale(locale)

    assert (
        render_file_admission_code(
            "FILE_EMPTY",
            declared_format="markdown",
            detected_format="unknown",
        )
        == expected
    )


def test_validate_files_renders_warning_in_active_locale(tmp_path: Path) -> None:
    init_cli_locale("de_DE")
    source = tmp_path / "image.jpg"
    source.write_bytes(b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR")

    valid, invalid, warnings = validate_files([str(source)])

    assert valid == [str(source.resolve())]
    assert invalid == []
    assert len(warnings) == 1
    assert "Der Dateiname gibt JPEG an" in warnings[0][1]
    assert "PNG" in warnings[0][1]
    assert "File extension" not in warnings[0][1]


def test_inspect_human_output_localizes_primary_warning_once(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    init_cli_locale("de_DE")
    source = tmp_path / "image.jpg"
    source.write_bytes(b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR")
    args = _build_parser().parse_args(["inspect", str(source)])

    assert execute_inspect(args) == 0

    output = capsys.readouterr().out
    assert output.count("Der Dateiname gibt JPEG an") == 1
    assert "[FILE_FORMAT_SAME_FAMILY_MISMATCH]" in output
    assert "File extension" not in output


def test_json_error_localizes_message_but_preserves_typed_admission(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    init_cli_locale("de_DE")
    source = tmp_path / "layout.docx"
    source.write_bytes(b"%PDF-1.4\n")
    output = tmp_path / "layout.md"
    args = _build_parser().parse_args(
        ["convert", str(source), "--to", "md", "--output", str(output), "--dry-run", "--json"]
    )

    assert execute_execution(args) == 2

    envelope = json.loads(capsys.readouterr().out)
    assert envelope["error"]["code"] == "file_format_confirmation_required"
    assert envelope["error"]["details"]["admission"]["warning_code"] == "FILE_FORMAT_CROSS_FAMILY_MISMATCH"
    assert envelope["error"]["details"]["admission"]["detected_format"] == "pdf"
    assert "Der Dateiname gibt DOCX an" in envelope["error"]["message"]
    assert "--use-detected-format" in envelope["error"]["hint"]


def test_json_warning_localizes_message_and_keeps_stable_code(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    init_cli_locale("de_DE")
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

    assert execute_execution(args, _available_controller()) == 0

    envelope = json.loads(capsys.readouterr().out)
    assert envelope["warnings"][0]["code"] == "FILE_FORMAT_CROSS_FAMILY_MISMATCH"
    assert envelope["warnings"][0]["details"] == {
        "decision": "require_explicit_acceptance",
        "declared_format": "docx",
        "detected_format": "pdf",
        "relation": "cross_family_mismatch",
    }
    assert "Der Dateiname gibt DOCX an" in envelope["warnings"][0]["message"]
