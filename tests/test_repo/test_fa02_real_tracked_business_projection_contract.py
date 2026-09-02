"""Fail-closed evidence guards for VIS-175 / finite-contract FA-02."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_docx_to_markdown_rich_semantics.json"
REPORT_NAME = "docx-real-tracked-business-final-artifact-parity-2026-07-22.md"
COMPLETE_PROJECTION_REPORT_NAME = "docx-real-tracked-business-complete-source-projection-2026-07-23.md"
REFERENCE_DISPOSITION_REPORT_NAME = "fa02-legacy-reference-defect-disposition-2026-07-23.md"
STATUS = "CURRENT_RESOURCE_FIX_VERIFIED_STRICT_ORACLE_RED_UNACCEPTED"


def _read(path: Path) -> str:
    return read_source_text(path)


def test_fa02_accepted_view_repairs_have_direct_guards() -> None:
    renderer = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "shared"
        / "markdown_runs.py"
    )
    textbox = _read(
        PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "docx_parsing" / "textbox_extraction.py"
    )
    breaks = _read(PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "docx_parsing" / "break_utils.py")
    parity_tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_to_markdown_standard_parity_*.py"
    )
    textbox_tests = _read(PROJECT_ROOT / "packages" / "core" / "tests" / "test_docx_textbox_extraction.py")

    assert 'elif tag in ("ins", "moveTo", "fldSimple", "smartTag", "sdt", "sdtContent", "customXml"):' in renderer
    assert 'elif tag == "noBreakHyphen":' in renderer
    assert "if vanish is not None and _on_off_property_is_enabled(vanish):" in renderer
    assert "processed_text_keys: set[tuple[int | None, str]]" in textbox
    assert 'if local_name in {"del", "moveFrom"}:' in textbox
    assert 'if local_name in {"del", "moveFrom"}:' in breaks
    assert "test_deleted_page_break_does_not_erase_accepted_inserted_paragraph" in parity_tests
    assert "test_drawingml_textbox_uses_accepted_revision_and_visible_run_projection" in textbox_tests


def test_fa02_resource_repair_has_direct_executable_guards() -> None:
    converter = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "to_markdown"
        / "converter.py"
    )
    parity_tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_to_markdown_standard_parity_*.py"
    )
    golden_tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_document_to_md_golden_*.py"
    )

    assert "for index in range(len(lines) - 1, -1, -1):" in converter
    assert "merged[index] += merged_refs" in converter
    assert "return [merged_refs, *lines]" in converter
    assert "test_image_refs_merge_once_after_last_nonempty_rendered_line" in parity_tests
    assert "test_image_only_element_still_emits_reachable_reference_line" in parity_tests
    assert "test_image_only_docx_paragraph_is_finalized_and_linked" in golden_tests
