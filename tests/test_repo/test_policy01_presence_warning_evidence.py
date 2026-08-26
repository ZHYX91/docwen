"""Fail-closed evidence guards for VIS-202 / POLICY-01=B."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "policy01-presence-warning-implementation-2026-07-23.md"
STAGE_CARD_NAME = "policy01-presence-warning-implementation-stage-card-2026-07-23.md"
FIXTURE_NAME = "current_policy01_presence_warning_semantics.json"
FIXTURE = ROOT / "tests/fixtures/golden" / FIXTURE_NAME


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_policy01_user_selection_and_accepted_boundary_are_exact() -> None:
    data = _fixture()
    assert data["status"] == "PASS_WITH_USER_ACCEPTED_BOUNDARY"
    assert data["user_selection"] == "POLICY-01=B"
    assert data["accepted_boundary"]["classification"] == "USER_ACCEPTED_BOUNDARY"
    statement = data["accepted_boundary"]["statement"]
    assert "cannot distinguish an intact signature from tampering" in statement
    assert "no signer, integrity, trust-chain, timestamp, or revocation assurance" in statement
    assert data["fixed_oracle"]["pass_threshold"] == "100%"
    assert data["fixed_oracle"]["pass"] is True


def test_fixed_corpus_invalid_siblings_and_external_result_are_frozen() -> None:
    data = _fixture()
    assert data["source_corpus"]["reused_stage"] == "VIS-2026-07-17-116"
    assert len(data["source_corpus"]["signed"]) == 3
    assert len(data["source_corpus"]["unsigned"]) == 3
    invalid = data["invalid_siblings"]
    assert invalid["entry_order_preserved"] is True
    assert invalid["all_other_entry_payloads_preserved"] is True
    assert invalid["signature_graph_payloads_preserved"] is True
    assert invalid["expected_structural_state"] == "complete"
    assert {invalid[owner]["mutated_part"] for owner in ("docx", "xlsx", "pptx")} == {
        "word/document.xml",
        "xl/sharedStrings.xml",
        "ppt/slides/slide1.xml",
    }
    evidence = data["external_evidence"]
    assert (evidence["files"], evidence["bytes"]) == (19, 268541)
    assert (evidence["result_bytes"], evidence["result_sha256"]) == (
        61588,
        "7aaf4845ea69560f07fba15faa41e0228c38758f766cc08b0d4d003480ebbc16",
    )
    assert evidence["binaries_checked_in"] is False


def test_all_conversion_inspect_and_gui_slots_pass_at_one_hundred_percent() -> None:
    oracle = _fixture()["fixed_oracle"]
    assert oracle["conversion_slots"] == {
        "expected": 15,
        "passed": 15,
        "owners_to_markdown": 9,
        "docx_xlsx_to_pdf": 6,
        "artifacts_present": 15,
    }
    assert oracle["inspect_slots"] == {
        "expected": 9,
        "passed": 9,
        "json_and_text": True,
        "unsigned_warning_free": True,
    }
    assert oracle["gui_slots"] == {
        "expected": 9,
        "passed": 9,
        "admission": True,
        "batch_admission": True,
        "durable_history": True,
    }
    contract = _fixture()["diagnostic_contract"]
    assert contract["signed_or_invalid_success_codes"] == [
        "OOXML_SIGNATURE_VALIDATION_UNAVAILABLE",
        "OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED",
    ]
    assert contract["unsigned_success_codes"] == []
    assert contract["unreadable_without_signature_markers_codes"] == []
    assert contract["failed_conversion_derived_output_code"] is False


def test_shared_detector_freezing_and_cross_surface_consumers_remain_wired() -> None:
    detector = (ROOT / "packages/core/src/docwen_core/detection/ooxml_signature.py").read_text(encoding="utf-8")
    controller = (ROOT / "packages/application/src/docwen_application/controller.py").read_text(encoding="utf-8")
    runtime = (ROOT / "packages/runtime/src/docwen_runtime/engine/task_manager.py").read_text(encoding="utf-8")
    cli = (ROOT / "packages/apps/cli/src/docwen_cli/commands/inspect.py").read_text(encoding="utf-8")
    inspection_model = (ROOT / "packages/core/src/docwen_core/models/file_inspection.py").read_text(encoding="utf-8")
    main_vm = (ROOT / "packages/apps/gui/src/docwen_gui/view_models/main_window_vm.py").read_text(encoding="utf-8")
    batch_vm = (ROOT / "packages/apps/gui/src/docwen_gui/view_models/batch_list_vm.py").read_text(encoding="utf-8")

    assert "inspect_ooxml_signature_graph" in detector
    assert "signature_validation_diagnostic" in detector
    assert "signature_derived_output_diagnostic" in detector
    assert "xmlsec" not in detector.lower()
    assert "cryptography" not in detector.lower()
    assert controller.index("freeze_ooxml_signature_info(request)") < controller.index("self._maybe_preconvert(request")
    assert "if result.success" in runtime
    assert "delivered_artifact=bool(result.artifacts)" in runtime
    assert "inspection.to_dict()" in cli
    assert "ooxml_signature" in inspection_model and "warnings" in inspection_model
    assert "inspect_file" in main_vm
    assert "inspect_file" in batch_vm
    assert "get_file_info" not in main_vm
    assert "get_file_info" not in batch_vm

    plugin_sources = "\n".join(
        path.read_text(encoding="utf-8", errors="replace") for path in (ROOT / "packages/plugins").rglob("*.py")
    ).lower()
    assert "_xmlsignatures" not in plugin_sources
    assert "origin.sigs" not in plugin_sources
