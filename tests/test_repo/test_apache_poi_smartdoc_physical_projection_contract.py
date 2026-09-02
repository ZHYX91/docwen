"""Fail-closed contracts for VIS-2026-07-17-113 physical SmartDoc evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIELD_FIXTURE = GOLDEN / "old_system_apache_poi_review_field_header_semantics.json"
ATTACHMENT_FIXTURE = GOLDEN / "old_system_apache_poi_attachment_revision_semantics.json"
REPORT_NAME = "apache-poi-smartdoc-physical-matrix-2026-07-17.md"


def _load(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def test_shared_physical_evidence_identity_and_execution_are_pinned() -> None:
    field = _load(FIELD_FIXTURE)["smartdoc_physical_addendum"]
    attachment = _load(ATTACHMENT_FIXTURE)["smartdoc_physical_addendum"]
    assert field["evidence_id"] == attachment["evidence_id"] == "VIS-2026-07-17-113"
    assert field["shared_external_evidence"] == attachment["shared_external_evidence"]
    evidence = field["shared_external_evidence"]
    assert (evidence["files"], evidence["bytes"]) == (133, 13_293_276)
    assert (evidence["production_artifacts"], evidence["word_pdfs"]) == (27, 30)
    assert (evidence["rendered_page_pngs"], evidence["contact_sheets"]) == (40, 9)
    assert evidence["physical_projection_sha256"] == (
        "ebdae3ef82a4d905e7d9ad695a49d77c039449c7c6e6c96bc90c0b27c65c533e"
    )
    assert evidence["word_object_projection_sha256"] == (
        "9843ae129a5561c27eabb56b5265c2e42841cd192471fde867005e55e4e0020f"
    )
    assert field["execution"]["production_conversions"] == 9
    assert attachment["execution"]["production_conversions"] == 18
    assert field["execution"]["all_contact_sheets_inspected"] is True


def test_all_nine_source_target_triples_are_exact_three_project_physical_matches() -> None:
    field = _load(FIELD_FIXTURE)["smartdoc_physical_addendum"]
    attachment = _load(ATTACHMENT_FIXTURE)["smartdoc_physical_addendum"]
    sources = {
        "FieldCodes.docx": field["FieldCodes.docx"],
        "delins.docx": attachment["delins.docx"],
        "WordWithAttachments.docx": attachment["WordWithAttachments.docx"],
    }
    page_counts = {"FieldCodes.docx": 1, "delins.docx": 1, "WordWithAttachments.docx": 2}
    for name, source in sources.items():
        assert set(source["same_target_three_project"]) == {"doc", "rtf", "odt"}
        for target in source["same_target_three_project"].values():
            assert target == {
                "pages": page_counts[name],
                "page_sizes_equal": True,
                "text_equal": True,
                "pixels_equal": True,
            }


def test_source_fidelity_defects_remain_shared_and_not_overclaimed() -> None:
    field = _load(FIELD_FIXTURE)["smartdoc_physical_addendum"]["FieldCodes.docx"]
    assert field["word_field_type_ids"] == {
        "source": [17, 21],
        "doc_all_projects": [17, 3],
        "rtf_all_projects": [17, 21],
        "odt_all_projects": [17, 21],
    }
    assert field["source_fidelity"]["doc"]["global_text_equal"] is False
    assert "not accepted broad field fidelity" in field["classification"]

    attachment = _load(ATTACHMENT_FIXTURE)["smartdoc_physical_addendum"]
    revisions = attachment["delins.docx"]
    assert revisions["source_fidelity"]["odt"]["changed_pixel_ratio"] == [0.089225]
    assert revisions["word_object_projection"]["odt_all_projects"] == {
        "paragraphs": 23,
        "fields": 18,
    }
    rich = attachment["WordWithAttachments.docx"]
    assert rich["source_fidelity"]["doc"]["global_pdf_text_equal"] is True
    assert rich["source_fidelity"]["rtf"]["text_sequence_ratio"] == 0.939612
    assert rich["source_fidelity"]["odt"]["text_sequence_ratio"] == 0.908371
    assert rich["word_object_counts_all_targets_all_projects"]["inline_shapes"] == 8
    assert "not attachment-byte extraction" in rich["classification"]
