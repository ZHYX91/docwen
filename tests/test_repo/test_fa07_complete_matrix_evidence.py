"""Fail-closed evidence guards for VIS-213 / finite-contract FA-07."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = [pytest.mark.contract, pytest.mark.golden]

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests" / "fixtures" / "golden" / "current_fa07_complete_matrix_semantics.json"
REPORT_NAME = "fa07-complete-matrix-artifact-oracle-2026-07-24.md"
STAGE_CARD = "fa07-complete-matrix-artifact-oracle-stage-card-2026-07-24.md"
DECISION_REPORT = "fa07-ods-fidelity-boundary-acceptance-2026-07-26.md"
DECISION_CARD = "fa07-ods-fidelity-boundary-acceptance-stage-card-2026-07-26.md"
ACQUISITION_REPORT = "fa07-ofgem-real-financial-model-acquisition-2026-07-24.md"
ACQUISITION_CARD = "fa07-ofgem-real-financial-model-acquisition-stage-card-2026-07-24.md"
STATUS = "PASS_WITH_USER_ACCEPTED_BOUNDARY"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_fa07_complete_matrix_fixture_accounts_for_exact_frozen_outputs() -> None:
    data = json.loads(_read(FIXTURE))

    assert data["stage_id"] == "VIS-2026-07-24-213"
    assert data["contract_id"] == "CONTRACT-2026-07-22-001/FA-07"
    assert data["sources"]["FA-07-B1"]["sha256"] == ("8066B12C878D74AA35B55B22DC9DABA9A56E91332A35CF96C4FC0F6AF6BD3BD1")
    assert data["sources"]["FA-07-N1"]["sha256"] == ("F284453E69F014A60C2732D2344BFE32EE114015784590D5BF729EF72CC42870")
    assert data["matrix"]["logical_slots"] == 24
    assert data["matrix"]["successful_slots"] == 20
    assert data["matrix"]["failed_or_blocked_slots"] == 4
    assert all(data["matrix"]["source_immutability"].values())
    assert data["matrix"]["current_n1_ods"]["returned_xlsx_sha256"] == (
        "B9D75F4A39E28B8426006AEE08921367428998AE1AD7DB7FBDD194B84D414261"
    )


def test_fa07_delivery_warning_and_user_accepted_boundary_remain_exact() -> None:
    data = json.loads(_read(FIXTURE))

    assert data["delivery_warning"]["diagnostic_codes"] == [
        "EXTERNAL_LINK_FLATTENED",
        "ODS_FEATURE_FIDELITY_RISK",
    ]
    assert data["delivery_warning"]["prepared_bytes_match_pre_warning_evidence"] is True
    assert data["delivery_warning"]["artifact_generation_reused"] is True
    assert data["delivery_warning"]["risk_counts"] == {
        "data_validations": 136,
        "conditional_formatting_ranges": 481,
        "charts": 31,
        "drawings": 34,
        "tables": 8,
        "pivot_or_slicer_parts": 10,
        "defined_names": 8735,
    }
    assert data["returned_xlsx_projection"]["current_n1_ods"]["data_validations"] == 0
    assert (
        data["returned_xlsx_projection"]["legacy_xls_fallback_is_more_destructive"]["direct_minus_fallback_formulas"]
        == 177406
    )
    assert data["physical_consumers"]["excel_returned_xlsx"]["opened"] == 10
    assert data["physical_consumers"]["excel_returned_xlsx"]["repair_or_open_errors"] == 0
    assert data["physical_consumers"]["wps_intermediates"]["timeouts"] == 3
    assert data["physical_consumers"]["libreoffice_intermediates"]["timeouts"] == 1
    assert data["visual_oracle"]["n1_current_ods_vs_source"]["exact_pages"] == 0
    assert data["visual_oracle"]["n1_current_ods_vs_source"]["common_pages"] == 11
    assert data["disposition"]["FA-07"] == STATUS
    assert data["disposition"]["decision_stage_id"] == "VIS-2026-07-26-385"
    assert data["disposition"]["user_decision_date"] == "2026-07-26"
    assert data["disposition"]["user_decision_quote"] == "那就接受"
    assert data["disposition"]["overall"] == "NOT PASSED YET"
