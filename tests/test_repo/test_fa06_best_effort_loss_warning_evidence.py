"""Fail-closed evidence guards for VIS-200 FA-06 best-effort delivery truth."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa06-best-effort-loss-warning-implementation-2026-07-23.md"
CARD_NAME = "fa06-best-effort-loss-warning-stage-card-2026-07-23.md"
STATUS = "BEST_EFFORT_WARNING_IMPLEMENTED_MATRIX_AND_ARTIFACT_ORACLE_PENDING"
CODE = "DOCX-SMARTDOC-BEST-EFFORT-LOSS"


def _read(path: Path) -> str:
    return read_source_text(path)


def _normalized(path: Path) -> str:
    return " ".join(_read(path).split())


def test_vis200_smartdoc_owner_delivers_typed_target_scoped_warning() -> None:
    source = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "to_document"
        / "converter.py"
    )
    tests = _read(PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_smart_converter.py")

    for token in (
        CODE,
        '"doc": "fields, revisions/comments, inline-object identities, and layout"',
        '"rtf": "fields, revisions/comments, inline-object identities, and layout"',
        '"odt": "paragraphs, tables, fields, revisions, shapes, sections, and pagination"',
        'request_source_format == "docx"',
        'source_format == "docx"',
        "target_format in _BEST_EFFORT_LOSS_CLASSES",
        'and "msoffice_word" in priority',
        'level="warning"',
        "source file was not modified",
        "Review the output against the source",
    ):
        assert token in source
    for token in (
        "test_docx_legacy_targets_deliver_with_typed_best_effort_warning",
        "test_best_effort_warning_is_limited_to_selected_docx_outbound_targets",
        "test_docx_to_rtf_prefers_word_then_keeps_configured_fallback_order",
        "test_docx_to_rtf_does_not_add_excluded_word_backend",
        "test_two_hop_route_does_not_apply_direct_docx_best_effort_preference",
        'assert "lossless" not in warnings[0].message.lower()',
        'read_bytes() == b"dummy binary content"',
    ):
        assert token in tests


def test_vis200_reuses_existing_cli_gui_success_warning_consumers() -> None:
    cli_text = _read(PROJECT_ROOT / "packages" / "apps" / "cli" / "tests" / "test_cli_text_diagnostics.py")
    cli_json = _read(PROJECT_ROOT / "packages" / "apps" / "cli" / "tests" / "test_cli_json.py")
    gui = _read(PROJECT_ROOT / "packages" / "apps" / "gui" / "tests" / "test_main_window_projection_binding_*.py")

    assert "test_single_success_writes_warning_diagnostic_to_stderr" in cli_text
    assert "test_batch_success_warning_includes_input_file" in cli_text
    assert "test_single_projects_only_warning_diagnostics" in cli_json
    assert "test_success_callback_projects_warning_diagnostics_to_info_area" in gui
    assert "test_batch_all_success_with_warning_keeps_success_state_and_warning_tone" in gui
