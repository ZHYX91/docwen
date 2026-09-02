from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_apache_poi_attachment_revision_semantics.json"


def _data() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_attachment_revision_sources_are_frozen_distinct_and_not_distributed() -> None:
    data = _data()
    repo = data["source_repository"]
    assert (repo["owner"], repo["repository"], repo["commit"]) == (
        "apache",
        "poi",
        "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96",
    )
    assert (repo["docx_paths_in_tree"], repo["unique_docx_blobs"]) == (128, 127)
    assert repo["package_screened_candidates_before_docwen_execution"] == 3
    expected = {
        "delins.docx": (17720, "7ecad102602586be7f1371f117fc567ea47f3ed0ca703b353e003002f1ff8d3f"),
        "EmbeddedDocument.docx": (13268, "3cb521da0faab3b8a9e7c5d155961b4f5506fb2ceda8e69e0790d03ee01767f2"),
        "WordWithAttachments.docx": (144959, "90352263823e6d3d39a4ec88212134b74bcb0f86b603bc3b51d712a3ec3df9f2"),
    }
    assert {name: (item["bytes"], item["sha256"]) for name, item in data["sources"].items()} == expected
    for name in expected:
        assert not any((ROOT / "tests/fixtures/files").rglob(name))


def test_revision_source_and_post_fix_policy_are_explicit() -> None:
    source = _data()["sources"]["delins.docx"]
    assert source["track_revisions"] is True
    assert (source["insertion_nodes"], source["insertion_bearing_paragraphs"], source["insertion_only_paragraphs"]) == (
        32,
        13,
        12,
    )
    assert (source["deletion_nodes"], source["deletion_only_paragraphs"], source["hyperlinks"]) == (4, 2, 7)
    projection = _data()["revision_projection"]
    assert projection["docwen-ref-tk"]["insertion_only_paragraphs_visible"] == 0
    assert projection["docwen-ref-pyside6"]["insertion_only_paragraphs_visible"] == 0
    assert projection["docwen-current-before-fix"]["insertion_only_paragraphs_visible"] == 0
    assert projection["docwen-current-after-fix"]["insertion_only_paragraphs_visible"] == 12
    assert {item["deletion_only_paragraphs_visible"] for item in projection.values() if isinstance(item, dict)} == {0}
    assert projection["docwen-current-after-fix"]["markdown_links"] == 7


def test_ole_sources_are_real_but_shared_omission_is_not_final_fidelity() -> None:
    sources = _data()["sources"]
    embedded = sources["EmbeddedDocument.docx"]
    assert len(embedded["embeddings"]) == len(embedded["preview_media"]) == 1
    assert embedded["embeddings"][0]["magic"] == "d0cf11e0a1b11ae1"
    multi = sources["WordWithAttachments.docx"]
    assert (multi["embedding_count"], multi["preview_media_count"], multi["hyperlinks"]) == (5, 5, 3)
    assert multi["textbox_paragraphs"] == ["Словарь", "dictionary", "Луғат"]
    projection = _data()["attachment_projection"]
    assert projection["artifact_count_per_project_per_source"] == 1
    assert projection["attachment_artifacts_emitted"] == 0
    assert projection["preview_image_artifacts_emitted"] == 0
    assert projection["embedded_filenames_emitted_in_markdown"] == 0
    assert projection["shared_omission_is_accepted_final_fidelity"] is False


def test_textbox_anchor_and_old_pyside_hyperlink_improvements_are_retained() -> None:
    item = _data()["attachment_projection"]["WordWithAttachments.docx"]
    assert item["docwen-ref-tk"]["hyperlinks_visible"] is False
    assert item["docwen-ref-pyside6"]["hyperlinks_visible"] is True
    assert item["docwen-current-after-fix"]["hyperlinks_visible"] is True
    assert item["docwen-current-before-fix"]["textbox_order"] == "document end"
    assert item["docwen-current-after-fix"]["textbox_order"] == "after source body anchor before table"
    classification = _data()["classification"]
    assert classification["old_pyside6_hyperlink_improvement_retained"] is True
    assert classification["source_order_textbox_improvement_retained"] is True


def test_current_runtime_finalizer_and_fix_classification_are_complete() -> None:
    data = _data()
    execution = data["execution_contract"]
    assert (execution["baseline_three_project_executions"], execution["post_fix_current_executions"]) == (9, 3)
    assert execution["valid_production_executions"] == 12
    assert execution["all_executions_successful"] is True
    runtime = data["current_runtime_finalizer"]
    assert runtime["artifact_count_per_source"] == 1
    assert runtime["diagnostics_per_source"] == ["DOCX2MD-OK", "FINALIZER_DONE"]
    assert runtime["workspace_paths_leaked"] is False
    assert runtime["delins.docx"]["paragraph_count"] == 21
    assert runtime["WordWithAttachments.docx"]["table_count"] == 1
    classification = data["classification"]
    assert classification["current_only_production_gap_found"] is True
    assert classification["existing_capability_but_real_entry_not_wired_found"] is True
    assert classification["documentation_or_test_evidence_drift_found"] is True
    assert classification["production_change_made"] is True
    assert data["fixes"]["layer_boundary_preserved"] is True
