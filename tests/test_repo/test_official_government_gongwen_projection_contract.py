from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_docx_official_government_list_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_official_gongwen_probe_reuses_the_hash_pinned_source_contract() -> None:
    data = _fixture()
    probe = data["gongwen_optimizer_probe"]

    assert data["related_golden_ids"] == ["GOLDEN-002", "GOLDEN-008"]
    assert probe["evidence_id"] == "VIS-2026-07-16-093"
    assert probe["source_reference"] == "$.source"
    assert data["source"]["bytes"] == 29246
    assert data["source"]["sha256"] == ("67d955f1ad1c71ca18221ac342093dce90c6d233ae9d48077afbc7ef6ad53a03")
    assert probe["execution_contract"]["external_binary_redistributed"] is False
    assert not (PROJECT_ROOT / "tests" / "fixtures" / "files" / "010941060n2v.docx").exists()


def test_official_gongwen_probe_records_three_project_missing_field_tolerance() -> None:
    projects = _fixture()["gongwen_optimizer_probe"]["projects"]
    for project in projects.values():
        assert project["success"] is True
        assert project["yaml_field_count"] == 18
        assert project["attachment_markdown_present"] is False

    assert projects["docwen-ref-tk"]["structured_missing_field_feedback_available"] is False
    for name in ("docwen-ref-pyside6", "docwen-current"):
        assert projects[name]["validation_status"] == "needs_review"
        assert projects[name]["missing_required"] == [
            "issue_date",
            "issuing_authority_signature",
        ]

    assert (
        projects["docwen-ref-tk"]["logical_markdown_sha256"]
        == projects["docwen-ref-pyside6"]["logical_markdown_sha256"]
    )
    assert projects["docwen-current"]["source_paragraphs_reachable_after_whitespace_normalization"] == 81


def test_official_gongwen_probe_locks_runtime_review_projection() -> None:
    probe = _fixture()["gongwen_optimizer_probe"]
    current = probe["projects"]["docwen-current"]
    fixed = probe["confirmed_current_gap_and_fix"]

    assert fixed["classification"] == "existing_pipeline_review_capability_not_wired_to_runtime_result"
    assert fixed["pre_fix_runtime_diagnostic_codes"] == ["GONGWEN-OK", "FINALIZER_DONE"]
    determinism = fixed["recognition_determinism"]
    assert determinism["verified_python_hash_seeds"] == [1, 999]
    assert determinism["official_sample_missing_required_at_both_seeds"] == [
        "issue_date",
        "issuing_authority_signature",
    ]
    assert current["diagnostic_codes"] == [
        "GONGWEN-OK",
        "GONGWEN-NEEDS-REVIEW",
        "FINALIZER_DONE",
    ]
    assert current["review_diagnostic_level"] == "warning"
    assert "成文日期、发文机关署名" in current["review_diagnostic_message"]
    assert current["artifact_metadata"]["gongwen_needs_review"] is True
    assert current["artifact_metadata"]["gongwen_missing_required"] == [
        "issue_date",
        "issuing_authority_signature",
    ]
    assert current["workspace_path_leaked"] is False
