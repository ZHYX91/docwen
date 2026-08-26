"""Phase A: Markdown -> DOCX -> Markdown numbering round-trip tests.

These are repo-level integration tests (see ``tests/integration``): the
chain crosses two plugin packages — ``docwen_plugin_markdown`` for the
forward (MD -> DOCX) leg and ``docwen_plugin_document`` for the reverse
(DOCX -> MD) leg — so they drive the full runtime orchestration rather
than deep-importing either plugin.

What "semantic stability" means here (the plan asks for semantic, not
byte-level, assertions):

* The reverse DOCX->MD leg **defaults to ``remove_numbering=True``
  internally**, which strips every heading prefix matched by the shared
  clean rules. That would mask whether forward-leg numbering survived,
  so the round-trip helper flips the reverse default to
  ``preserve_numbering=True`` (sends ``remove_numbering=False``). The
  survival tests below rely on this; a separate test locks down the
  default-stripping behaviour when preservation is explicitly off.

* **text render mode** — added numbering is baked into heading *text*
  (e.g. ``一、引言``). With the reverse leg preserving, it survives the
  round-trip as literal text and must reappear. Assert the prefixes
  are present.

* **word_native render mode** — added numbering is a Word multi-level
  *list structure* in the forward DOCX; heading text stays bare there.
  The reverse DOCX->MD leg reconstructs numbering text from that Word
  list structure, so the round-tripped Markdown headings carry the
  scheme prefixes (the structure was honored end-to-end). Assert the
  headings come back with numbering and **without** doubled or garbled
  prefixes — i.e. the round-trip is stable, not that it comes back
  bare.

* **handwritten input + remove + add** — input Markdown already
  carries handwritten numbering. With the shared clean rules loaded
  (as the session fixture does in production parity), the forward
  ``remove_numbering`` strips the handwritten prefix and
  ``add_numbering`` re-applies a clean one, yielding **single**
  numbering across the round-trip. Assert no doubled prefixes.

* **4-scheme parametrized** — the text-mode survival contract must
  hold for every system scheme, so each scheme gets its own case.
"""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.integration

from tests.integration._round_trip_helper import round_trip_md

# ── Shared fixtures ───────────────────────────────────────────────────────

_PLAIN_MD = "# 引言\n\n正文一。\n\n## 背景\n\n正文二。\n\n## 方法\n\n正文三。\n"

_HANDWRITTEN_MD = "# 一、引言\n\n正文一。\n\n## （一）背景\n\n正文二。\n\n## （二）方法\n\n正文三。\n"


def _write_md(tmp_path: Path, name: str, content: str) -> Path:
    """Write a Markdown input file into tmp_path and return its path."""
    p = tmp_path / name
    p.write_text(content, encoding="utf-8")
    return p


def _headings(md_text: str) -> list[str]:
    """Return the text of every ATX heading line, stripped of leading ``#``."""
    out: list[str] = []
    for line in md_text.splitlines():
        stripped = line.strip()
        if stripped.startswith("#"):
            # count leading hashes
            level = 0
            for ch in stripped:
                if ch == "#":
                    level += 1
                else:
                    break
            text = stripped[level:].strip()
            if text:
                out.append(text)
    return out


# ── 1. text mode minimal round-trip ──────────────────────────────────────


class TestTextModeRoundTrip:
    """Step 1: ``text`` render mode — added numbering survives as literal text."""

    def test_added_numbering_survives_round_trip(self, round_trip_runtime, tmp_path: Path) -> None:
        """text + gongwen_standard: forward-added prefixes reappear in reverse MD."""
        md_input = _write_md(tmp_path, "plain.md", _PLAIN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "add_numbering": True,
                "heading_numbering_render_mode": "text",
                "numbering_scheme": "gongwen_standard",
            },
        )

        headings = _headings(md_text)
        assert headings, f"No headings found in round-trip output:\n{md_text[:500]}"

        # The h1 should carry the gongwen level-1 prefix (一、).
        assert any(h.startswith("一、") and "引言" in h for h in headings), (
            f"h1 should carry '一、' prefix after text-mode add. Headings: {headings}\nOutput:\n{md_text[:500]}"
        )

        # The two h2 headings should carry the gongwen level-2 prefix （一）/（二）.
        h2s = [h for h in headings if h.startswith(("（一）", "（二）"))]
        assert len(h2s) >= 2, (
            f"Expected >=2 h2 headings with （一）/（二） prefixes. Headings: {headings}\nOutput:\n{md_text[:500]}"
        )

    def test_heading_structure_preserved(self, round_trip_runtime, tmp_path: Path) -> None:
        """text mode: heading levels and body text survive the round-trip."""
        md_input = _write_md(tmp_path, "plain.md", _PLAIN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "add_numbering": True,
                "heading_numbering_render_mode": "text",
                "numbering_scheme": "gongwen_standard",
            },
        )

        # Body paragraphs must survive.
        assert "正文一" in md_text, f"Body text '正文一' lost. Output:\n{md_text[:500]}"
        assert "正文二" in md_text, f"Body text '正文二' lost. Output:\n{md_text[:500]}"
        assert "正文三" in md_text, f"Body text '正文三' lost. Output:\n{md_text[:500]}"

        # Heading level markers survive (one h1, two h2).
        assert md_text.count("\n# ") + md_text.startswith("# ") >= 1
        assert md_text.count("\n## ") >= 2 or "## " in md_text

    def test_default_reverse_leg_strips_numbering(self, round_trip_runtime, tmp_path: Path) -> None:
        """The reverse DOCX->MD leg defaults to remove_numbering=True.

        Locks down the cross-layer fact: when preservation is off
        (the converter's internal default), every prefix matched by
        the shared clean rules is stripped from the round-tripped
        Markdown. This is the behaviour GUI/CLI users get by default
        on DOCX->MD, and it is why the survival tests must explicitly
        preserve.
        """
        md_input = _write_md(tmp_path, "plain.md", _PLAIN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "add_numbering": True,
                "heading_numbering_render_mode": "text",
                "numbering_scheme": "gongwen_standard",
            },
            preserve_numbering=False,
        )

        headings = _headings(md_text)
        # With stripping on, the gongwen prefixes (一、 / （一）) must be gone.
        for h in headings:
            assert not h.startswith("一、"), f"prefix '一、' should be stripped by default reverse leg: {h!r}"
            assert not h.startswith("（一）"), f"prefix '（一）' should be stripped by default reverse leg: {h!r}"
        # Bare heading words must remain.
        assert any("引言" in h for h in headings), f"bare heading text should remain after strip. Headings: {headings}"


# ── 2. word_native mode round-trip ───────────────────────────────────────


class TestWordNativeRoundTrip:
    """Step 2: ``word_native`` render mode — numbering is structural in DOCX.

    The forward leg applies numbering as a Word multi-level list, so
    heading *text* stays bare in the DOCX. The reverse DOCX->MD leg
    reconstructs numbering text from that list structure, so the
    round-tripped Markdown headings carry the scheme prefixes. The
    contract is therefore: the round-trip is **stable** — headings
    come back with single, scheme-correct numbering, no doubling, no
    garbled prefixes, and body text intact.
    """

    def test_headings_come_back_with_numbering(self, round_trip_runtime, tmp_path: Path) -> None:
        """word_native: reverse leg reconstructs numbering; h1 carries 一、."""
        md_input = _write_md(tmp_path, "plain.md", _PLAIN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "add_numbering": True,
                "heading_numbering_render_mode": "word_native",
                "numbering_scheme": "gongwen_standard",
            },
        )

        headings = _headings(md_text)
        assert headings, f"No headings found:\n{md_text[:500]}"

        # h1 should carry the reconstructed level-1 prefix.
        assert any(h.startswith("一、") and "引言" in h for h in headings), (
            f"word_native h1 should carry reconstructed '一、' prefix. Headings: {headings}\nOutput:\n{md_text[:500]}"
        )

        # Bare heading words must still be present (not eaten by the list).
        assert any("引言" in h for h in headings), f"heading word '引言' lost. Headings: {headings}"
        assert any("背景" in h for h in headings), f"heading word '背景' lost. Headings: {headings}"

    def test_no_double_numbering_in_output(self, round_trip_runtime, tmp_path: Path) -> None:
        """word_native: no heading text is doubled or garbled by the round-trip."""
        md_input = _write_md(tmp_path, "plain.md", _PLAIN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "add_numbering": True,
                "heading_numbering_render_mode": "word_native",
                "numbering_scheme": "gongwen_standard",
            },
        )

        # No heading should contain a repeated prefix or doubled word.
        for h in _headings(md_text):
            assert "一、一、" not in h, f"doubled level-1 prefix: {h!r}"
            assert "（一）（一）" not in h, f"doubled level-2 prefix: {h!r}"
            assert "引言引言" not in h, f"doubled heading text: {h!r}"
            assert "背景背景" not in h, f"doubled heading text: {h!r}"

    def test_body_text_survives(self, round_trip_runtime, tmp_path: Path) -> None:
        """word_native: body paragraphs survive the structural-numbering round-trip."""
        md_input = _write_md(tmp_path, "plain.md", _PLAIN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "add_numbering": True,
                "heading_numbering_render_mode": "word_native",
                "numbering_scheme": "gongwen_standard",
            },
        )

        assert "正文一" in md_text, f"body '正文一' lost. Output:\n{md_text[:500]}"
        assert "正文二" in md_text, f"body '正文二' lost. Output:\n{md_text[:500]}"
        assert "正文三" in md_text, f"body '正文三' lost. Output:\n{md_text[:500]}"


# ── 3. handwritten input + remove/add ────────────────────────────────────


class TestHandwrittenRemoveAddRoundTrip:
    """Step 3: input already has handwritten numbering; remove+add should
    yield *single* numbering, not double.

    With the shared clean rules loaded (as the session fixture does,
    matching production parity), the forward ``remove_numbering``
    strips the handwritten Chinese prefixes (``一、`` / ``（一）``)
    before ``add_numbering`` re-applies clean ones. The round-trip
    therefore surfaces single, scheme-correct numbering.
    """

    def test_remove_then_add_yields_single_numbering(self, round_trip_runtime, tmp_path: Path) -> None:
        """remove+add on handwritten input must not double the numbering."""
        md_input = _write_md(tmp_path, "hand.md", _HANDWRITTEN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "remove_numbering": True,
                "add_numbering": True,
                "heading_numbering_render_mode": "text",
                "numbering_scheme": "gongwen_standard",
            },
        )

        headings = _headings(md_text)
        # No heading should carry a doubled prefix.
        for h in headings:
            assert "一、一、" not in h, f"doubled level-1 prefix: {h!r}"
            assert "（一）（一）" not in h, f"doubled level-2 prefix: {h!r}"
            assert "（二）（二）" not in h, f"doubled level-2 prefix: {h!r}"

        # And the clean scheme prefix should be present (single).
        assert any(h.startswith("一、") and "引言" in h for h in headings), (
            f"single '一、' prefix should be present after remove+add. Headings: {headings}"
        )

    def test_handwritten_numbering_preserved_without_remove(self, round_trip_runtime, tmp_path: Path) -> None:
        """Without remove, handwritten numbering is preserved across the round-trip.

        This documents current (correct) behaviour: when the user does
        not request removal, handwritten prefixes pass through to DOCX
        heading text and back. This is the faithful baseline against
        which the remove+add behaviour above is measured.
        """
        md_input = _write_md(tmp_path, "hand.md", _HANDWRITTEN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "add_numbering": False,
                "remove_numbering": False,
            },
        )

        headings = _headings(md_text)
        assert any(h.startswith("一、") and "引言" in h for h in headings), (
            f"handwritten '一、' prefix should be preserved when remove is off. Headings: {headings}"
        )
        assert any(h.startswith("（一）") and "背景" in h for h in headings), (
            f"handwritten '（一）' prefix should be preserved. Headings: {headings}"
        )


# ── 4. four-scheme parametrized (text mode) ──────────────────────────────

_SCHEME_IDS = [
    "gongwen_standard",
    "hierarchical_standard",
    "hierarchical_h2_start",
    "legal_standard",
]


# Per-scheme expected prefix shape on the round-tripped h2 headings
# (h1 is bare by design for hierarchical_h2_start). These are the
# prefixes the forward text-mode add bakes into heading text and that
# the preserving reverse leg must surface intact. A scheme-specific
# regression (silent no-op, wrong format, prefix lost) breaks the
# matching prefix for that scheme only.
_SCHEME_H2_PREFIX_EXPECTED: dict[str, str] = {
    "gongwen_standard": "（",  # （一）背景 / （二）方法
    "hierarchical_standard": "1.1",  # 1.1 背景 / 1.2 方法
    "hierarchical_h2_start": "1 ",  # 1 背景 / 2 方法  (h1 bare by design)
    "legal_standard": "第一章",  # 第一章　背景 / 第二章　方法
}


class TestFourSchemeRoundTrip:
    """Step 4: the text-mode survival contract holds for every system scheme."""

    @pytest.mark.parametrize("scheme_id", _SCHEME_IDS)
    def test_added_numbering_survives_for_each_scheme(self, round_trip_runtime, tmp_path: Path, scheme_id: str) -> None:
        """text mode + each scheme: forward-added numbering reappears in reverse MD.

        The exact prefix differs per scheme; the test asserts each
        scheme's distinctive h2 prefix is present in the round-tripped
        output. This catches a scheme-specific regression where add
        silently no-ops or the prefix is lost for one scheme but not
        others.
        """
        md_input = _write_md(tmp_path, f"{scheme_id}.md", _PLAIN_MD)
        work = tmp_path / "work"

        _docx, md_text = round_trip_md(
            round_trip_runtime,
            md_input,
            work,
            forward_options={
                "add_numbering": True,
                "heading_numbering_render_mode": "text",
                "numbering_scheme": scheme_id,
            },
        )

        headings = _headings(md_text)
        assert headings, f"[{scheme_id}] no headings in round-trip output:\n{md_text[:500]}"

        expected_prefix = _SCHEME_H2_PREFIX_EXPECTED[scheme_id]
        h2s = [h for h in headings if "背景" in h or "方法" in h]
        assert h2s, f"[{scheme_id}] h2 headings (背景/方法) not found. Headings: {headings}"
        assert any(h.startswith(expected_prefix) for h in h2s), (
            f"[{scheme_id}] no h2 carries expected prefix {expected_prefix!r}. "
            f"h2 headings: {h2s}\nFull output:\n{md_text[:500]}"
        )

        # Body text must always survive regardless of scheme.
        assert "正文一" in md_text, f"[{scheme_id}] body text '正文一' lost. Output:\n{md_text[:500]}"
