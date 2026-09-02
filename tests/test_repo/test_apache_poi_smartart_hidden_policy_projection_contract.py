"""Fail-closed contracts for VIS-2026-07-17-131 Presentation policy evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_pptx_to_markdown_semantics.json"
REPORT_NAME = "apache-poi-smartart-hidden-slide-policy-2026-07-17.md"


def _addendum() -> dict[str, object]:
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    return fixture["smartart_hidden_policy_addendum"]


def test_source_powerpoint_oracle_and_package_order_are_pinned() -> None:
    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-131"
    assert addendum["source_commit"] == "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96"
    assert addendum["source"] == {
        "filename": "60810.pptx",
        "git_blob": "3e4e22be5f4911be6d9935246d32f5f76907ba4f",
        "bytes": 874_522,
        "sha256": "61afceb0365523ceba6fc00a525157148c2ddbdb3b5d843992ac3c107e9921bf",
    }

    oracle = addendum["powerpoint_source_oracle"]
    assert (oracle["consumer"], oracle["version"], oracle["build"]) == (
        "Microsoft PowerPoint",
        "16.0",
        "20131",
    )
    assert oracle["slide_count"] == 28
    assert oracle["hidden_slide_indexes"] == [4, 21, 22]
    assert oracle["smartart_slide_indexes"] == [8, 13, 22]
    assert oracle["smartart_diagram_count"] == 4
    assert oracle["smartart_text_count"] == 34
    assert oracle["smartart_text_count_by_slide"] == {"8": 7, "13": 11, "22": 16}
    assert oracle["package_order_exact_after_com_encoding_normalization"] is True


def test_pre_fix_shared_omission_is_not_reclassified_as_accepted() -> None:
    matrix = _addendum()["pre_fix_three_project_matrix"]
    assert matrix["success_count"] == matrix["expected_success_count"] == 3
    assert matrix["old_markdown_raw_exact"] is True
    assert matrix["smartart_bullet_count"] == {
        "docwen-ref-tk": 0,
        "docwen-ref-pyside6": 0,
        "docwen-current-pre": 0,
    }
    assert matrix["classification"] == "shared_unaccepted_source_fidelity_gap"


def test_current_enhancement_is_exact_and_preserves_prior_artifacts() -> None:
    addendum = _addendum()
    current = addendum["current_enhancement"]
    assert current["classification"] == "current_enhancement_closes_shared_unaccepted_boundary"
    assert current["diagnostic_codes"] == ["PPTX2MD-OK", "FINALIZER_DONE"]
    assert current["primary_metadata"] == {
        "slide_count": 28,
        "hidden_slide_count": 3,
        "table_count": 0,
        "image_count": 19,
        "smartart_text_count": 34,
        "title": "3.7.1 HRBP Process overview",
    }
    assert current["smartart_bullet_count"] == 34
    assert current["smartart_bullet_order_exact"] is True
    assert current["base_nonempty_lines_unchanged"] is True
    assert current["current_resource_names_hashes_exact"] is True
    assert current["three_project_resource_hash_multisets_exact"] is True

    production = (ROOT / current["production"]).read_text(encoding="utf-8")
    regression = (ROOT / current["regression"]).read_text(encoding="utf-8")
    assert "def _extract_smartart_texts(" in production
    assert 'value in {"0", "false", "off", "no"}' in production
    assert 'lines.extend(f"- {text}" for text in smartart_texts)' in production
    assert "test_process_slide_extracts_smartart_nodes_in_relationship_order" in regression
    assert "test_parse_pptx_includes_hidden_slide_content_and_counts_policy" in regression


def test_hidden_slide_policy_and_environment_boundary_are_explicit() -> None:
    addendum = _addendum()
    policy = addendum["hidden_slide_policy"]
    assert policy["policy"] == "include_all_source_slides_and_report_hidden_count"
    assert "not silently discarded" in policy["reason"]
    assert policy["config_or_request_option_added"] is False

    process = addendum["process_boundary"]
    assert process["powerpoint_pid_present_immediately_after_quit"] == 20668
    assert process["powerpoint_present_post_settlement"] is False
    assert process["post_settlement_only_preexisting_wps_pid"] == 11388
    assert process["harness_termination_command_used"] is False

    classification = addendum["classification"]
    assert classification == {
        "current_only_regression_found": False,
        "smartart_content": "closed_for_pinned_real_source",
        "hidden_slide_policy": "explicit_and_regression_tested",
        "broad_presentation_or_overall_parity": "not_proven",
    }
