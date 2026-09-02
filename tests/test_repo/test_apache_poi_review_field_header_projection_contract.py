from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_apache_poi_review_field_header_semantics.json"


def _data() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_poi_docx_sources_are_frozen_distinct_and_not_distributed() -> None:
    data = _data()
    repo = data["source_repository"]
    assert (repo["owner"], repo["repository"], repo["commit"]) == (
        "apache",
        "poi",
        "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96",
    )
    assert (repo["docx_paths_in_tree"], repo["unique_docx_blobs"]) == (128, 127)
    assert repo["tree_truncated"] is False
    expected = {
        "FieldCodes.docx": (17034, "9247f85ab90b15f11d0cf14d0070fdcb5cdf47c65d6bc77cc32cb7ee682aad24"),
        "headerFooter.docx": (28423, "0ad90c6ee8ec8b9fab1ae8219387eca8d934b739580918d76a1b6c911bc41913"),
        "headerPic.docx": (16206, "50d9b17f7575c8a91129fd8a3e6a70feec886480ab9f089e29a7f8ac7e5ed51a"),
        "testComment.docx": (65298, "06b63b13f0b78a7179151af52c2282557c1ed204899feee9fcef4ca575344f97"),
    }
    assert {name: (item["bytes"], item["sha256"]) for name, item in data["sources"].items()} == expected
    for name in expected:
        assert not any((ROOT / "tests/fixtures/files").rglob(name))


def test_poi_docx_package_features_cover_complex_field_review_and_headers() -> None:
    sources = _data()["sources"]
    field = sources["FieldCodes.docx"]
    assert (field["field_chars"], field["instruction_text_nodes"], field["simple_fields"]) == (6, 2, 0)
    assert field["saved_display_values"] == ["ANTONI", "16 June 2010"]
    comment = sources["testComment.docx"]
    assert (
        comment["comment_count"],
        comment["comment_range_start_count"],
        comment["comment_range_end_count"],
        comment["comment_reference_count"],
    ) == (1, 1, 1, 1)
    assert comment["comment_text"] == "comment content"
    assert comment["media"]["relationship_owner"] == "word/comments.xml"
    assert sources["headerFooter.docx"]["nonempty_header_text"].endswith("…")
    assert sources["headerPic.docx"]["media"]["relationship_owner"] == "word/header1.xml"


def test_poi_docx_twelve_production_runs_have_equal_normalized_projections() -> None:
    data = _data()
    execution = data["execution_contract"]
    assert execution["valid_three_project_executions"] == 12
    assert execution["all_executions_successful"] is True
    expected = {
        "FieldCodes.docx": (3, "2ad95b91e24f211e3ec9e992900a6579390b314744e470522bd5c9ca3761bd75"),
        "headerFooter.docx": (0, "0a8801ac16afb402e1f93d1c57db68f08126cc8a7348d8642816d130bbba14dc"),
        "headerPic.docx": (0, "b689c3007762e0f867a643012659b6b1ed2fdef031f45e0eed125fec8b6b1a14"),
        "testComment.docx": (1, "e047450a2fc60f95ebc3f0e36c415e3ac34f1971d958f5563b591d67b2cd5857"),
    }
    for name, (line_count, digest) in expected.items():
        item = data["normalized_projections"][name]
        assert item["all_three_projects_equal"] is True
        assert item["tk_old_pyside6_raw_equal"] is True
        assert len(item["body_lines"]) == line_count
        assert item["canonical_sha256"] == digest
    assert "CRLF" in data["normalized_projections"]["current_raw_difference"]


def test_review_field_and_header_outputs_keep_shared_boundary_without_overclaim() -> None:
    feature = _data()["feature_projection"]
    assert feature["complex_field_saved_values_visible_in_all_projects"] is True
    assert feature["comment_anchor_body_visible_in_all_projects"] is True
    for key in (
        "complex_field_instruction_text_emitted",
        "comment_text_emitted",
        "comment_owned_image_emitted",
        "header_footer_text_emitted",
        "header_owned_image_emitted",
    ):
        assert feature[key] is False
    classification = _data()["classification"]
    assert classification["current_only_functional_gap_found"] is False
    assert classification["shared_omission_is_accepted_final_fidelity"] is False


def test_poi_docx_current_runtime_finalizer_is_complete_for_all_sources() -> None:
    runtime = _data()["current_runtime_finalizer"]
    assert runtime["artifact_count_per_source"] == 1
    assert runtime["diagnostics_per_source"] == ["DOCX2MD-OK", "FINALIZER_DONE"]
    assert runtime["workspace_paths_leaked"] is False
    expected = {
        "FieldCodes.docx": ("FieldCodes.md", 95, 2, 0),
        "headerFooter.docx": ("headerFooter.md", 77, 0, 0),
        "headerPic.docx": ("headerPic.md", 71, 0, 0),
        "testComment.docx": ("testComment.md", 103, 1, 0),
    }
    for name, facts in expected.items():
        item = runtime[name]
        assert item["success"] is True
        assert (item["primary_name"], item["primary_bytes"], item["paragraph_count"], item["image_count"]) == facts
