"""Fail-closed evidence guards for VIS-201 / selected FA-08 closure."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
REPORT_NAME = "fa08-delivery-first-source-fidelity-implementation-2026-07-23.md"
STAGE_CARD = "fa08-delivery-first-source-fidelity-stage-card-2026-07-23.md"
STATUS = "FIXED_AND_VERIFIED_WITH_USER_ACCEPTED_MPO_AUXILIARY_PAGE_BOUNDARY"
WARNING_CODE = "IMG2PDF-MPO-AUXILIARY-FRAMES"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def test_selected_policy_is_implemented_by_the_owned_image_routes() -> None:
    common = _read("packages/plugins/image/src/docwen_plugin_image/_common.py")
    flat = _read("packages/plugins/image/src/docwen_plugin_image/format_conversion/converter.py")
    pdf = _read("packages/plugins/image/src/docwen_plugin_image/to_pdf/converter.py")
    tests = _read("packages/plugins/image/tests/test_image_conversions_*.py")

    for token in (
        "def prepare_flat_export(",
        "ImageOps.exif_transpose(img)",
        "FLAT_EXPORT_EXIF_TAGS = (271, 272, 305, 306)",
        'save_metadata["icc_profile"] = icc_profile',
    ):
        assert token in common
    assert 'if target in ("jpg", "webp"):' in flat
    assert "prepare_flat_export(img)" in flat
    assert 'if img.format == "MPO":' in pdf
    assert WARNING_CODE in pdf
    for selector in (
        "test_fa08_flat_export_normalizes_orientation_and_preserves_supported_metadata",
        "test_fa08_mpo_delivers_every_frame_with_auxiliary_warning",
        "test_fa08_single_frame_pdf_does_not_emit_mpo_warning",
    ):
        assert selector in tests


def test_post_choice_oracle_and_fixture_lock_the_exact_boundary() -> None:
    evaluator = _read("tools/validation/evaluate_fa08_delivery_first.py")
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    closure = fixture["fa08_delivery_first_closure_addendum"]
    historical = fixture["fa08_final_artifact_contract_addendum"]

    for token in (
        'EXPECTED_STAGE_CONTRACT_SHA256 = "c1299d3a4e76007ef3ad21efecf8b9ad1773cb55562ebdfee39c652bfd7f8909"',
        "SUPPORTED_EXIF = {271, 272, 274, 305, 306}",
        "output_icc == source_icc",
        "N1_ACCEPTED_AUXILIARY_SCORE = 0.79880938",
        "affected_slots_passed",
    ):
        assert token in evaluator

    assert historical["status"] == "pending_user_decision_shared_source_fidelity"
    assert historical["pass"] is False
    assert closure["policy"] == "FA-08=O-A,E-A,C-A,M-B"
    assert closure["status"] == STATUS.lower()
    assert closure["affected_current_slots"] == 8
    assert closure["affected_current_slots_passed"] == 8
    assert closure["flat_export_contract"]["icc_payload_exact"] is True
    assert closure["mpo_pdf_contract"]["warning_code"] == WARNING_CODE
    assert closure["mpo_pdf_contract"]["n1_auxiliary_page_render_score"] == 0.79880938
    assert closure["pass"] is True
