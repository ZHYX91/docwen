from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from docx import Document

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_docx_multilingual_template_batch_semantics.json"

EXPECTED_SOURCE_HASHES = {
    "Deutsche Allgemeine Vorlage.docx": "4c13650734dc0697111336a61fcf51402a34565f2a879ce2b2beec3e39ca6e05",
    "English General Template.docx": "ee5483b3c8229e41ca88c63fd02027c47679851204ac776d14c5941e8d131e04",
    "Modelo Geral Português.docx": "87d9b580955a45606e2c00658e8937a6b030b928fffa14dc3a391067a92e0600",
    "Modèle Général Français.docx": "6bf38c14b0ee8491c85fe0d74b0859c9e21a4eae9376b251fad7e32548acbc2f",
    "Mẫu Chung Tiếng Việt.docx": "a53b164ab4a40a7c6232a4c0a6a3962c5d0f8155359e97c3eba83879afab455c",
    "Plantilla General Española.docx": "c679d482ed6b8ff39b1c1fd2afcab1b7ed34f85f8796fd60199ccdd83ea12ca5",
    "Русский Общий Шаблон.docx": "bb60955407e23fde6b2cdf052903b8fc49cae25e1293e41ff85c514050762a74",
    "日本語汎用テンプレート.docx": "d59e15471ca1139228fd558379f94dc63be5279d23066dd8ae71865e90107eb5",
    "简体中文通用模板.docx": "399a2d1337b326f97bdab390f9e465446a26f88a3e2329602eaa7c961460c3b9",
    "繁體中文通用模板.docx": "cb3bd3962ea64dfd45a456a3f59f39269d54735ee51e40b105a08ef48f2bb64b",
    "한국어 범용 템플릿.docx": "d06cce71c06045b8c9c53996deedf39b8175b71908f8b6dca98866bc84b79720",
}


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_multilingual_template_fixture_has_exact_real_asset_inventory() -> None:
    data = _fixture()
    contract = data["input_contract"]
    templates = contract["templates"]

    assert data["golden_id"] == "GOLDEN-002"
    assert data["evidence_id"] == "VIS-2026-07-15-027"
    assert contract["template_count"] == 11
    assert contract["templates_are_checked_in_distribution_assets"] is True
    assert contract["source_bytes_are_identical_across_all_three_projects"] is True
    assert contract["project_order"] == [
        "docwen-ref-tk",
        "docwen-ref-pyside6",
        "docwen-current",
    ]
    assert {item["name"] for item in templates} == set(EXPECTED_SOURCE_HASHES)


def test_multilingual_template_fixture_hashes_and_localized_styles_match_assets() -> None:
    templates = {item["name"]: item for item in _fixture()["input_contract"]["templates"]}
    for name, expected_hash in EXPECTED_SOURCE_HASHES.items():
        source = PROJECT_ROOT / "templates" / name
        assert source.is_file()
        assert hashlib.sha256(source.read_bytes()).hexdigest() == expected_hash
        assert templates[name]["source_sha256"] == expected_hash
        assert templates[name]["source_bytes"] == source.stat().st_size

        document = Document(source)
        paragraphs = [paragraph for paragraph in document.paragraphs if paragraph.text.strip()]
        assert templates[name]["styles"] == [
            paragraphs[0].style.name,
            paragraphs[1].style.name,
        ]
        assert templates[name]["tokens"] == [paragraphs[0].text, paragraphs[1].text]


def test_multilingual_template_fixture_locks_normalized_three_project_results() -> None:
    data = _fixture()
    normalized = data["normalized_contract"]
    assert normalized["all_eleven_yaml_semantics_equal"] is True
    assert normalized["all_eleven_frontmatter_values_equal"] is True
    assert normalized["all_eleven_body_lines_equal"] is True
    assert normalized["all_eleven_placeholder_token_sets_equal"] is True
    assert normalized["no_unicode_replacement_characters"] is True
    assert normalized["old_tk_and_old_pyside6_raw_markdown_equal_for_all_templates"] is True
    assert normalized["raw_current_byte_equality_required"] is False
    assert "double-quoted" in normalized["raw_current_difference"]

    for item in data["input_contract"]["templates"]:
        old_tk, old_pyside6, current = item["markdown_sha256"]
        old_tk_bytes, old_pyside6_bytes, current_bytes = item["markdown_utf8_bytes"]
        assert old_tk == old_pyside6
        assert current != old_tk
        assert old_tk_bytes == old_pyside6_bytes
        assert current_bytes == old_tk_bytes + 1
        assert item["current_primary_name"] == Path(item["name"]).with_suffix(".md").name


def test_multilingual_template_fixture_locks_execution_finalizer_and_harness_boundaries() -> None:
    data = _fixture()
    execution = data["execution_contract"]
    current = data["current_runtime_contract"]

    assert execution["valid_project_template_executions"] == 33
    assert execution["all_valid_executions_successful"] is True
    assert "translator injection" in execution["docwen_ref_tk_entry"]
    assert "initialize_runtime" in execution["docwen_ref_pyside6_entry"]
    assert "TaskManager -> OutputFinalizer" in execution["docwen_current_entry"]
    assert len(execution["excluded_harness_attempts"]) == 3
    assert any("0/11" in item for item in execution["excluded_harness_attempts"])
    assert any("_001" in item for item in execution["excluded_harness_attempts"])

    assert current["all_primary_names_preserve_unicode_source_stems"] is True
    assert current["all_primary_artifacts_in_requested_output_directories"] is True
    assert current["all_have_docx2md_and_finalizer_diagnostics"] is True
    assert current["all_metadata"] == {
        "paragraph_count": 1,
        "heading_count": 0,
        "table_count": 0,
        "image_count": 0,
    }
    assert data["process_evidence"]["new_after_final_settle"] == []


def test_multilingual_template_evidence_updates_actual_golden_inventory_only() -> None:
    golden_files = sorted((PROJECT_ROOT / "tests" / "fixtures" / "golden").glob("*.json"))
    assert len(golden_files) == 85
    assert FIXTURE in golden_files
    for path in golden_files:
        json.loads(path.read_text(encoding="utf-8"))
