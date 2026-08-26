"""Golden baseline test: compare new MD→DOCX output against old-system golden.

This test:
1. Copies ``samples/sample.md`` to a temp directory.
2. Runs the new CLI MD→DOCX conversion via ``docwen_bundle.cli_entry.main``.
3. Compares the generated DOCX against the old-system golden at
   ``tests/fixtures/golden/md_to_docx_old/sample_golden.docx``.
4. Asserts on the comparison result.

When the golden does NOT match, the test fails with a detailed diff so the
developer can decide whether the change is expected (golden needs updating)
or a regression (code needs fixing).
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

pytestmark = [pytest.mark.golden, pytest.mark.contract]
from tools.validation.compare_md_to_docx_golden import (
    _compare_headings,
    _compare_paragraphs,
)

# ── Paths relative to repo root ────────────────────────────────────────

REPO_ROOT = Path(__file__).resolve().parents[2]
SAMPLE_MD = REPO_ROOT / "samples" / "sample.md"
GOLDEN_DOCX = REPO_ROOT / "tests" / "fixtures" / "golden" / "md_to_docx_old" / "sample_golden.docx"
DOCX_TEMPLATE = REPO_ROOT / "templates" / "简体中文通用模板.docx"
COMPARE_TOOL = REPO_ROOT / "tools" / "validation" / "compare_md_to_docx_golden.py"


@pytest.mark.golden
class TestMdToDocxOldBaseline:
    """Golden comparison of new CLI MD→DOCX output vs. old-system baseline."""

    def test_matches_old_golden(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """Run MD→DOCX via CLI, then compare against old golden."""
        monkeypatch.setenv("DOCWEN_CONFIG_DIR", str(tmp_path / "user-config"))
        # 1. Copy sample.md into tmp_path so the CLI outputs there too.
        md_input = tmp_path / "sample.md"
        shutil.copy2(str(SAMPLE_MD), str(md_input))

        # 2. Run the CLI conversion.
        from docwen_bundle.cli_entry import main as cli_main
        from docwen_runtime.templates import TemplateRegistry

        new_docx = tmp_path / "sample.docx"
        template = next(
            item
            for item in TemplateRegistry.default().list_templates("docx")
            if item.path.resolve() == DOCX_TEMPLATE.resolve()
        )
        exit_code = cli_main(
            [
                "convert",
                str(md_input),
                "--to",
                "docx",
                "--output",
                str(new_docx),
                "--template",
                template.id,
            ]
        )
        assert exit_code == 0, f"CLI conversion failed with exit code {exit_code}"

        # Protocol 3 writes only to the explicit output path.
        assert new_docx.exists(), (
            f"Expected output DOCX not found at {new_docx}\ntmp_path contents: {list(tmp_path.iterdir())}"
        )

        # 3. Compare against the old golden.
        from tools.validation.compare_md_to_docx_golden import compare_docx_files

        result = compare_docx_files(
            str(GOLDEN_DOCX),
            str(new_docx),
            allowed_new_title="Test File",
            allowed_removed_old_paragraph="$$",
            allowed_heading_body_merge=(
                2,
                "Title Ending with Punctuation:",
                "Title Ending with Punctuation: When a heading ends with punctuation and is immediately followed by "
                "body text (no blank line), they form a combined heading-paragraph block.",
            ),
        )
        assert result["details"]["paragraphs"]["allowed_removed_old_paragraph_applied"] is True
        assert result["details"]["paragraphs"]["allowed_heading_body_merge_applied"] is True
        assert result["details"]["headings"]["allowed_heading_body_merge_applied"] is True

        # 4. Assert — with a rich failure message showing the diff details.
        if not result["passed"]:
            import json

            details = json.dumps(result["details"], ensure_ascii=False, indent=2)
            msg = (
                f"\n{'=' * 70}\n"
                f"GOLDEN MISMATCH: new CLI output differs from old baseline.\n"
                f"{'=' * 70}\n"
                f"  paragraphs_match : {result['paragraphs_match']}\n"
                f"  headings_match   : {result['headings_match']}\n"
                f"  tables_match     : {result['tables_match']}\n"
                f"\n"
                f"Old paragraphs: {result['old_paragraph_count']}  "
                f"New paragraphs: {result['new_paragraph_count']}\n"
                f"Old headings:   {result['old_heading_count']}  "
                f"New headings:   {result['new_heading_count']}\n"
                f"Old tables:     {result['old_table_count']}  "
                f"New tables:     {result['new_table_count']}\n"
                f"\nDiff details:\n{details}\n"
                f"{'=' * 70}"
            )
            pytest.fail(msg)

        # If we reach here, all three categories match.
        assert result["passed"] is True


def test_historical_math_delimiter_allowance_is_exact_and_unique() -> None:
    old = [
        {"style": "Normal", "text": "before"},
        {"style": "Normal", "text": "$$"},
        {"style": "Normal", "text": "after"},
    ]
    new = [
        {"style": "Normal", "text": "before"},
        {"style": "Normal", "text": "after"},
    ]

    accepted = _compare_paragraphs(old, new, allowed_removed_old_paragraph="$$")
    assert accepted["match"] is True
    assert accepted["allowed_removed_old_paragraph_applied"] is True

    duplicate_old = [old[0], old[1], old[1], old[2]]
    duplicate_result = _compare_paragraphs(duplicate_old, new, allowed_removed_old_paragraph="$$")
    assert duplicate_result["match"] is False
    assert duplicate_result["allowed_removed_old_paragraph_applied"] is False

    reordered_new = [new[1], new[0]]
    reordered_result = _compare_paragraphs(old, reordered_new, allowed_removed_old_paragraph="$$")
    assert reordered_result["match"] is False
    assert reordered_result["allowed_removed_old_paragraph_applied"] is True


def test_historical_heading_body_merge_allowance_is_exact_and_unique() -> None:
    allowance = (2, "Planning:", "Planning: Body text.")
    old_headings = [{"level": 2, "text": "Planning:", "style": "Heading 2"}]
    new_headings = [{"level": 2, "text": "Planning: Body text.", "style": "Heading 2"}]
    old_paragraphs = [
        {"style": "Heading 2", "text": "Planning:"},
        {"style": "Normal", "text": "Body text."},
    ]
    new_paragraphs = [{"style": "Heading 2", "text": "Planning: Body text."}]

    heading_result = _compare_headings(
        old_headings,
        new_headings,
        allowed_heading_body_merge=allowance,
    )
    paragraph_result = _compare_paragraphs(
        old_paragraphs,
        new_paragraphs,
        allowed_heading_body_merge=allowance,
    )

    assert heading_result["match"] is True
    assert heading_result["allowed_heading_body_merge_applied"] is True
    assert paragraph_result["match"] is True
    assert paragraph_result["allowed_heading_body_merge_applied"] is True

    duplicate_body_result = _compare_paragraphs(
        [*old_paragraphs, {"style": "Normal", "text": "Body text."}],
        new_paragraphs,
        allowed_heading_body_merge=allowance,
    )
    assert duplicate_body_result["match"] is False
    assert duplicate_body_result["allowed_heading_body_merge_applied"] is False
