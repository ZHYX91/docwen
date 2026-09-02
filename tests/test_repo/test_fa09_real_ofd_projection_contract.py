"""Fail-closed evidence guards for VIS-177 / finite-contract FA-09."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa09-real-ofd-n1-shared-reference-red-2026-07-22.md"
STAGE_CARD = "fa09-real-ofd-scanned-pdf-n1-stage-card-2026-07-22.md"
STATUS = "CURRENT_N1_PDF_FIXED_SHARED_REFERENCE_RED_UNACCEPTED"
SOURCE_SHA = "55192F12F6AFF1294EB9F40000F9455B5F905D5F661AA308EBFAFE8A31154D02"
OUTPUT_SHA = "72DA9011BAF1CD7887149D8B96AECED70218A494F3E8C84D0981021D67BA9391"


def _read(path: Path) -> str:
    return read_source_text(path)


def test_fa09_easyofd_repairs_have_direct_executable_guards() -> None:
    patches = _read(PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "ofd.py")
    converter = _read(
        PROJECT_ROOT / "packages" / "plugins" / "layout" / "src" / "docwen_plugin_layout" / "to_pdf" / "converter.py"
    )
    core_tests = _read(PROJECT_ROOT / "packages" / "core" / "tests" / "test_ofd_ocr_helpers.py")
    layout_tests = _read(PROJECT_ROOT / "packages" / "plugins" / "layout" / "tests" / "test_layout_conversions_*.py")

    assert "def _patch_content_clip_boundary" in patches
    assert "def _patch_draw_pdf_page_scale" in patches
    assert "self.OP = 72 / 25.4" in patches
    assert "apply_easyofd_patches()" in converter
    assert "test_unbounded_abbreviated_clip_does_not_crash_text_parser" in core_tests
    assert "test_bounded_clip_retains_upstream_position_parsing" in core_tests
    assert "test_draw_pdf_uses_pdf_points_per_millimetre" in core_tests
    assert "test_ofd_to_pdf_applies_easyofd_patches" in layout_tests
