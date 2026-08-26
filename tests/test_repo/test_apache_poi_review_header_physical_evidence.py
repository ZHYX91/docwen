"""Fail-closed contracts for VIS-2026-07-17-121 review/header evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_apache_poi_review_field_header_semantics.json"
REPORT_NAME = "apache-poi-review-header-physical-matrix-2026-07-17.md"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def _addendum() -> dict[str, object]:
    return _fixture()["review_header_physical_addendum"]


def test_source_identity_and_external_projection_are_pinned() -> None:
    fixture = _fixture()
    assert fixture["source_repository"]["commit"] == ("86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96")
    expected_sources = {
        "headerFooter.docx": (
            28_423,
            "0ad90c6ee8ec8b9fab1ae8219387eca8d934b739580918d76a1b6c911bc41913",
        ),
        "headerPic.docx": (
            16_206,
            "50d9b17f7575c8a91129fd8a3e6a70feec886480ab9f089e29a7f8ac7e5ed51a",
        ),
        "testComment.docx": (
            65_298,
            "06b63b13f0b78a7179151af52c2282557c1ed204899feee9fcef4ca575344f97",
        ),
    }
    for name, (size, digest) in expected_sources.items():
        source = fixture["sources"][name]
        assert (source["bytes"], source["sha256"]) == (size, digest)

    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-121"
    evidence = addendum["external_evidence"]
    assert evidence["inventory_excluding_harness_and_summary"] == {
        "files": 183,
        "bytes": 4_584_889,
    }
    assert evidence["root_including_harness_and_summary"] == {
        "files": 187,
        "bytes": 4_621_920,
    }
    assert evidence["evidence_summary"] == {
        "bytes": 7_155,
        "sha256": "09b7c020bfdb3e25c1698ca66a9302d60f8b20aaaf7ed5f139e74b66e326999b",
    }
    assert evidence["key_file_sha256"] == {
        "conversion-docwen-ref-tk.json": ("1d7502bc237c1f2479dc76cefbb23e853840680ac13bee481227a4dfda900b05"),
        "conversion-docwen-ref-pyside6.json": ("45201b4544626f4ea838b8d223398fcc89a4c61f769956b329b45b0662715b07"),
        "conversion-docwen-current.json": ("b14efdbfd817a05bf66f2560cf46446d39dc2b6a4583647d75f828e546127c1a"),
        "process-boundary.json": ("f445009b1ba8173432cc92cd9ce4d501d970ccd31aebbd71f3a2e68ea8dc210e"),
        "word-object-projection.json": ("56151b55b62d72b2111c3c15b4386d3f72e88faa5a2b87ba6402e62687662f02"),
        "word-render-process-boundary.json": ("604c4db1df8cd73677de7d0ae229c922a9f746f8e802bc4f7877707971e515fb"),
        "physical-projection.json": ("1b3a0e69907183781a414eefe83b93d1690a1b08069e7f6f0a505fc162d404f4"),
    }


def test_all_twelve_same_target_triples_are_exact() -> None:
    execution = _addendum()["execution"]
    assert execution["sources"] == [
        "headerFooter.docx",
        "headerPic.docx",
        "testComment.docx",
    ]
    assert execution["targets"] == ["doc", "rtf", "odt"]
    assert execution["production_conversions"] == 27
    assert execution["all_successful"] is True
    assert (
        execution["word_documents_opened"],
        execution["word_pdf_exports"],
        execution["rendered_views"],
        execution["rendered_pages"],
        execution["contact_sheets"],
    ) == (30, 40, 40, 40, 12)
    assert execution["all_contact_sheets_inspected"] is True

    same_target = _addendum()["same_target_three_project"]
    assert (same_target["triples"], same_target["content_triples"]) == (12, 9)
    assert same_target["comment_markup_triples"] == 3
    for key in (
        "page_count_equal",
        "render_size_equal",
        "visible_text_equal",
        "pixels_equal",
    ):
        assert same_target[key] is True


def test_header_comment_objects_and_physical_media_boundary_are_explicit() -> None:
    addendum = _addendum()
    assert addendum["word_object_preservation"] == {
        "header_footer_outputs_with_expected_text": 9,
        "header_picture_outputs_with_one_header_inline_shape": 9,
        "comment_outputs_with_one_comment_and_expected_text": 9,
    }
    observation = addendum["manual_physical_observation"]
    assert observation["header_footer_visible_in_all_outputs"] is True
    assert observation["header_picture_visible_in_all_outputs"] is True
    assert observation["comment_anchor_and_text_visible_in_all_markup_outputs"] is True
    assert observation["comment_owned_picture_visible_in_doc_outputs"] is True
    assert observation["comment_owned_picture_visible_in_rtf_outputs"] is True
    assert observation["comment_owned_picture_visible_in_odt_outputs"] is False
    assert observation["current_only_clip_overlap_black_block_missing_page_or_displacement"] is False

    classification = addendum["classification"]
    assert classification == {
        "current_only_functional_or_physical_regression_found": False,
        "production_change_made": False,
        "shared_odt_comment_owned_picture_loss": True,
        "shared_odt_comment_owned_picture_loss_is_accepted_final_fidelity": False,
        "broad_smartdoc_or_review_ui_pass": False,
    }
