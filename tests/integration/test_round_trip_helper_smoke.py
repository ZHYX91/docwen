"""Smoke test: validate the round-trip helper scaffolding itself.

This is not a numbering-semantics test (those live in
``test_md_docx_md_numbering_round_trip.py``). It exists only to prove
the cross-plugin orchestration helper wires up correctly: a trivial
Markdown document survives ``MD -> DOCX -> MD`` and comes back as
non-empty Markdown with its heading text preserved.

If this smoke test fails, the scaffolding is broken and every Phase A
test built on it would also fail for infrastructure reasons rather
than real regressions.
"""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.integration

from tests.integration._round_trip_helper import round_trip_md


def test_round_trip_helper_smoke(round_trip_runtime, tmp_path: Path) -> None:
    """A trivial MD doc round-trips through DOCX and back to non-empty MD."""
    md_input = tmp_path / "smoke.md"
    md_input.write_text(
        "# Smoke Title\n\nFirst paragraph.\n\n## Smoke Section\n\nSecond paragraph.\n",
        encoding="utf-8",
    )

    work = tmp_path / "work"
    docx_path, md_text = round_trip_md(round_trip_runtime, md_input, work)

    # The forward leg produced a real DOCX on disk.
    assert docx_path.exists(), f"DOCX not produced: {docx_path}"
    assert docx_path.suffix == ".docx"
    assert docx_path.stat().st_size > 0

    # The reverse leg produced non-empty Markdown.
    assert md_text.strip(), "Round-tripped Markdown is empty"

    # Heading text must survive the round-trip (semantic, not byte-level).
    assert "Smoke Title" in md_text, f"h1 heading text lost in round-trip. Output:\n{md_text[:500]}"
    assert "Smoke Section" in md_text, f"h2 heading text lost in round-trip. Output:\n{md_text[:500]}"
