"""Integration coverage for the local noncandidate 60-row DOCX rehearsal."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from scripts.release import verify_neutral_docx_60row as rehearsal

from docwen_core.docx_semantics import DocxSemanticImporter
from docwen_core.models.semantic_document import (
    SemanticCitationCluster,
    SemanticParagraph,
    SemanticReference,
    SemanticTable,
    SemanticText,
)

pytestmark = pytest.mark.integration

REPO_ROOT = Path(__file__).resolve().parents[2]


def _output(tmp_path: Path, suffix: str) -> Path:
    return tmp_path / f"{rehearsal.ARTIFACT_CLASS}-{suffix}"


def _all_strings(value: object) -> tuple[str, ...]:
    if isinstance(value, str):
        return (value,)
    if isinstance(value, dict):
        strings: list[str] = []
        for key, item in value.items():
            strings.extend(_all_strings(key))
            strings.extend(_all_strings(item))
        return tuple(strings)
    if isinstance(value, list):
        return tuple(item for child in value for item in _all_strings(child))
    return ()


def test_cli_generation_path_emits_headless_evidence_without_viewer_pass(tmp_path, capsys) -> None:
    output_dir = _output(tmp_path, "cli")

    assert rehearsal.main(["--output-dir", str(output_dir)]) == 0

    captured = capsys.readouterr()
    assert captured.err == ""
    cli_payload = json.loads(captured.out)
    assert cli_payload["artifact_class"] == rehearsal.ARTIFACT_CLASS
    docx_path = output_dir / rehearsal.DOCX_NAME
    report_path = output_dir / rehearsal.REPORT_NAME
    checklist_path = output_dir / rehearsal.MANUAL_CHECKLIST_NAME
    assert docx_path.is_file()
    assert report_path.is_file()
    assert checklist_path.is_file()

    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["schema"] == rehearsal.REPORT_SCHEMA
    assert report["artifact_class"] == rehearsal.ARTIFACT_CLASS
    assert report["source"]["head"]
    assert report["source"]["tree"]
    assert "branch" not in report["source"]
    assert isinstance(report["source"]["remotes"], list)
    assert report["source"]["status"] in {"CLEAN", "DIRTY"}
    assert report["source"]["porcelain"] == _git_porcelain_lines(REPO_ROOT)
    expected_overall = (
        "SOURCE_HEADLESS_PASS" if report["source"]["status"] == "CLEAN" else "SOURCE_HEADLESS_VERIFIED_DIRTY_SOURCE"
    )
    assert report["overall"] == expected_overall
    assert cli_payload["overall"] == expected_overall
    assert report["artifacts"]["docx"]["sha256"] == rehearsal._sha256(docx_path)
    assert report["artifacts"]["manual_checklist"]["sha256"] == rehearsal._sha256(checklist_path)
    assert report["automatic_gates"]
    assert all(report["automatic_gates"].values())
    assert report["evidence_layers"]["viewer_word"] == "READY_FOR_MANUAL"
    assert report["evidence_layers"]["viewer_wps"] == "READY_FOR_MANUAL"
    assert report["evidence_layers"]["viewer_libreoffice"] == "READY_FOR_MANUAL"
    assert not any(value in {"VIEWER_PASS", "CANDIDATE_PASS", "RELEASE_PASS"} for value in _all_strings(report))

    checklist = checklist_path.read_text(encoding="utf-8")
    assert "Microsoft Word" in checklist
    assert "WPS Writer" in checklist
    assert "LibreOffice Writer" in checklist
    assert "Pagination" in checklist
    assert "Two-row repeated header" in checklist
    assert "Horizontal Sales merge and vertical Region header merge" in checklist
    assert "Caption and cross-reference display the same number" in checklist
    assert "SEQ/REF source cache starts at 60" in checklist
    assert "viewer may recalculate the first table" in checklist
    assert "citation display" in checklist
    assert "no bibliography entries" in checklist
    assert checklist.count("Status: READY_FOR_MANUAL") == 3
    assert not any(token in checklist for token in ("VIEWER_PASS", "CANDIDATE_PASS", "RELEASE_PASS"))

    semantic_document = rehearsal.build_neutral_document_60row()
    verified = rehearsal.verify_docx_60row(docx_path, semantic_document)
    assert verified
    assert all(verified.values())


def test_libreoffice_style_field_rewrite_preserves_table_reference_and_fails_closed_citations(tmp_path: Path) -> None:
    semantic_document = rehearsal.build_neutral_document_60row()
    expected_table = semantic_document.blocks[0]
    assert isinstance(expected_table, SemanticTable)
    source = tmp_path / "neutral-60-row-source.docx"
    localized = tmp_path / "neutral-60-row-localized-seq.docx"
    rehearsal._render_and_save(semantic_document, source)

    viewer_document = Document(source)
    instruction = next(
        item
        for item in viewer_document.element.iter(qn("w:instrText"))
        if (item.text or "").strip().startswith("SEQ Table")
    )
    instruction.text = (instruction.text or "").replace("Table", "表格")
    locked_field_begins = [
        item
        for item in viewer_document.element.iter(qn("w:fldChar"))
        if item.get(qn("w:fldCharType")) == "begin" and item.get(qn("w:fldLock")) in {"1", "true", "on"}
    ]
    assert len(locked_field_begins) == 2
    for field_begin in locked_field_begins:
        field_begin.attrib.pop(qn("w:fldLock"))
    viewer_document.save(localized)

    reopened = Document(localized)
    imported = DocxSemanticImporter().import_document(reopened)
    imported_tables = [block for block in imported.document.blocks if isinstance(block, SemanticTable)]
    imported_references = [
        inline
        for block in imported.document.blocks
        if isinstance(block, SemanticParagraph)
        for inline in block.inlines
        if isinstance(inline, SemanticReference)
    ]

    assert "SEQ 表格" in reopened.element.xml
    assert "SEQ Table" not in reopened.element.xml
    assert [item.code for item in imported.diagnostics] == [
        "semantic.docx.citation.field_unlocked",
        "semantic.docx.citation.field_unlocked",
    ]
    assert len(imported_tables) == 1
    assert imported_tables[0].caption == expected_table.caption
    assert imported_references == [SemanticReference("tbl-neutral-60", "60")]
    assert SemanticParagraph((SemanticText("Citations [1] and [2, 1]."),)) in imported.document.blocks
    assert not any(
        isinstance(inline, SemanticCitationCluster)
        for block in imported.document.blocks
        if isinstance(block, SemanticParagraph)
        for inline in block.inlines
    )


def _git_porcelain_lines(repo: Path) -> list[str]:
    completed = subprocess.run(
        ["git", "-C", str(repo), "status", "--porcelain=v1", "--untracked-files=all"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return completed.stdout.rstrip("\n").splitlines()


@pytest.mark.parametrize(
    ("output_dir", "error_code"),
    (
        (Path(f"{rehearsal.ARTIFACT_CLASS}-relative"), "output_dir_must_be_absolute"),
        (REPO_ROOT / f"{rehearsal.ARTIFACT_CLASS}-inside-repo", "output_dir_inside_source_repo"),
    ),
)
def test_output_path_rejections_fail_before_creating_directory(output_dir: Path, error_code: str) -> None:
    assert not output_dir.exists()

    with pytest.raises(rehearsal.RehearsalError, match=error_code):
        rehearsal.run_rehearsal(output_dir, source_repo=REPO_ROOT)

    assert not output_dir.exists()


def test_preexisting_output_fails_closed_without_overwrite(tmp_path) -> None:
    output_dir = _output(tmp_path, "preexisting")
    output_dir.mkdir()
    sentinel = output_dir / "keep.txt"
    sentinel.write_text("unchanged", encoding="utf-8")

    with pytest.raises(rehearsal.RehearsalError, match="output_dir_exists"):
        rehearsal.run_rehearsal(output_dir, source_repo=REPO_ROOT)

    assert sentinel.read_text(encoding="utf-8") == "unchanged"
    assert list(output_dir.iterdir()) == [sentinel]


def test_output_directory_name_requires_noncandidate_marker(tmp_path) -> None:
    output_dir = tmp_path / "neutral-60-row-output"

    with pytest.raises(rehearsal.RehearsalError, match="output_dir_missing_noncandidate_marker"):
        rehearsal.run_rehearsal(output_dir, source_repo=REPO_ROOT)

    assert not output_dir.exists()
