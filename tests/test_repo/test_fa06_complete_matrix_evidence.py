"""Fail-closed evidence guards for VIS-206 / finite-contract FA-06 closure."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = [pytest.mark.contract, pytest.mark.golden]
ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa06-complete-matrix-artifact-oracle-2026-07-24.md"
CARD_NAME = "fa06-complete-matrix-artifact-oracle-stage-card-2026-07-24.md"
FIXTURE_NAME = "current_fa06_best_effort_complete_matrix_semantics.json"
FIXTURE = ROOT / "tests/fixtures/golden" / FIXTURE_NAME
STATUS = "PASS_WITH_USER_ACCEPTED_BOUNDARY"
WARNING = "DOCX-SMARTDOC-BEST-EFFORT-LOSS"


def _data() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_fa06_complete_matrix_identity_accounting_and_disposition_are_exact() -> None:
    data = _data()
    assert data["status"] == STATUS
    assert data["overall_parity"] == "NOT_PASSED_YET"
    assert data["user_selection"] == "FA-06=DOC-B,RTF-B,ODT-B"
    assert data["stage_start"] == {
        "parent": "6328cbdd52dcc5d091d0240f6d2bb4d40cb857b9",
        "tree": "6ca63f8037a2a46b8b0a63a3eef84e6908e35ecd",
    }
    assert data["frozen_sources"] == {
        "FA-06-B1": {
            "bytes": 29246,
            "sha256": "67d955f1ad1c71ca18221ac342093dce90c6d233ae9d48077afbc7ef6ad53a03",
        },
        "FA-06-N1": {
            "bytes": 151733,
            "sha256": "a27f2b9244f5147cce25a05d4cddb1eec9c72259a41d52a7bf5eb3806b28207c",
        },
    }
    accounting = data["execution_accounting"]
    assert accounting["logical_slots"] == 36
    assert accounting["reused_reference_outbound"] == 12
    assert accounting["first_pass_fresh"] == 24
    assert accounting["previously_unexecuted_completed"] == 21
    assert accounting["current_n1_outbound_revalidated"] == 3
    assert accounting["repair_pass_fresh"] == 12
    assert accounting["reference_inbound_reused_from_first_pass"] == 12
    assert accounting["final_outbound_artifacts"] == 18
    assert accounting["final_inbound_artifacts"] == 18
    assert accounting["pass_threshold"].startswith("100%")


def test_fa06_current_only_revision_red_is_fixed_without_hiding_source_revisions() -> None:
    data = _data()
    red = data["initial_current_only_red"]
    assert red == {
        "classification": "CONVERTER_CREATED_TRACKED_REVISIONS",
        "source_revision_count": 9,
        "current_doc_outbound_revision_count": 12,
        "current_rtf_outbound_revision_count": 11,
        "unexpected_author": "Yuxiang Zheng",
        "disposition": "FIXED_AND_REAL_ARTIFACT_VERIFIED",
    }
    fix = data["fix"]
    assert fix["source_revisions_preserved"] is True
    assert fix["generic_bridge_default_unchanged"] is True
    assert fix["current_n1_doc_revision_count_after_fix"] == 9
    assert fix["current_n1_rtf_revision_count_after_fix"] == 9

    bridge = (ROOT / "packages/core/src/docwen_core/office_bridge.py").read_text(encoding="utf-8")
    converter = (ROOT / "packages/plugins/document/src/docwen_plugin_document/to_document/converter.py").read_text(
        encoding="utf-8"
    )
    assert "suppress_new_revisions: bool = False" in bridge
    assert "doc_or_wb.TrackRevisions = False" in bridge
    assert converter.count("suppress_new_revisions=True") == 3


def test_fa06_mandatory_artifact_and_projection_predicates_are_all_satisfied() -> None:
    data = _data()
    predicates = data["mandatory_predicates"]
    assert predicates == {
        "all_36_slots_accounted": True,
        "all_artifacts_nonempty_and_container_valid": True,
        "source_immutable": True,
        "reused_reference_outbound_immutable": True,
        "current_outbound_warning_code": WARNING,
        "current_outbound_warning_count": 6,
        "conversion_process_delta": 0,
        "projection_process_delta": 0,
        "word_documents_opened_without_repair_or_save": 26,
        "pdf_exports": 39,
        "rendered_pages": 363,
        "contact_sheets_manually_inspected": 9,
        "current_only_regression_remaining": False,
    }
    assert data["object_projection"]["n1_source"]["revisions"] == 9
    assert data["object_projection"]["n1_current_doc_inbound"]["revisions"] == 9
    assert data["object_projection"]["n1_current_rtf_inbound"]["pages"] == 7
    assert data["object_projection"]["n1_current_odt_inbound"] == {
        "pages": 7,
        "paragraphs": 158,
        "tables": 2,
        "fields": 11,
        "comments": 1,
        "revisions": 0,
        "inline_shapes": 5,
        "shapes": 9,
        "sections": 1,
    }


def test_fa06_accepted_boundary_is_specific_and_objective_failures_stay_errors() -> None:
    boundary = _data()["accepted_boundary"]
    assert boundary["classification"] == "USER_ACCEPTED_BOUNDARY"
    assert "least-divergent valid artifact" in boundary["delivery_rule"]
    assert "ODT may expand the 13-page source to 14 pages" in boundary["b1"][1]
    assert "20-to-11 fields" not in " ".join(boundary["n1"])
    assert "reduce paragraphs, tables, fields, revisions and sections" in boundary["n1"][2]
    assert set(boundary["not_accepted"]) == {
        "backend failure",
        "timeout",
        "unreadable input",
        "invalid or empty output",
        "source mutation",
        "undisclosed loss",
        "converter-created tracked revisions",
    }
