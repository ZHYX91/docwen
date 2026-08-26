"""Tests for docwen_core.links error semantics — centralised error-kind
classification, placeholder generation, keep-link synthesis, and mode-aware
error-output dispatch.

Covers F-H2-010, F-H2-011, F-H2-012, F-H2-013, F-H2-014.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_core.links import (
    LinkErrorKind,
    NotFoundAction,
    dispatch_error_output,
    make_error_placeholder,
    make_keep_link,
    resolve_embedded_links,
)

pytestmark = pytest.mark.unit


# ═══════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════


def _write(path: Path, text: str) -> str:
    """Write *text* to *path* and return the absolute path string."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return str(path)


# ═══════════════════════════════════════════════════════════════════════════
# LinkErrorKind enum
# ═══════════════════════════════════════════════════════════════════════════


class TestLinkErrorKind:
    """Verify the enumeration has exactly the five error kinds."""

    def test_all_kinds_present(self) -> None:
        names = {k.name for k in LinkErrorKind}
        assert names == {
            "SECTION_NOT_FOUND",
            "BLOCK_NOT_FOUND",
            "FILE_NOT_FOUND",
            "CIRCULAR_REFERENCE",
            "MAX_DEPTH",
        }

    def test_values_are_distinct(self) -> None:
        vals = [k.value for k in LinkErrorKind]
        assert len(vals) == len(set(vals))

    def test_label_returns_chinese_text(self) -> None:
        assert LinkErrorKind.SECTION_NOT_FOUND.label == "章节未找到"
        assert LinkErrorKind.BLOCK_NOT_FOUND.label == "块未找到"
        assert LinkErrorKind.FILE_NOT_FOUND.label == "文件未找到"
        assert LinkErrorKind.CIRCULAR_REFERENCE.label == "循环引用"
        assert LinkErrorKind.MAX_DEPTH.label == "达到最大深度"

    def test_int_construction_normalised(self) -> None:
        assert LinkErrorKind(1) == LinkErrorKind.SECTION_NOT_FOUND
        assert LinkErrorKind(5) == LinkErrorKind.MAX_DEPTH

    def test_invalid_int_raises_value_error(self) -> None:
        with pytest.raises(ValueError, match="not a valid LinkErrorKind"):
            LinkErrorKind(99)


# ═══════════════════════════════════════════════════════════════════════════
# make_error_placeholder  (F-H2-010 … F-H2-014)
# ═══════════════════════════════════════════════════════════════════════════


class TestMakeErrorPlaceholder:
    """Covers F-H2-010 through F-H2-014: differentiated placeholder text
    for every error kind."""

    # ── section not found (F-H2-010) ────────────────────────────────────

    def test_section_not_found(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.SECTION_NOT_FOUND,
            "doc.md",
            heading="Intro",
        )
        assert result == "[Section not found: doc.md#Intro]"

    def test_section_not_found_no_heading(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.SECTION_NOT_FOUND,
            "doc.md",
        )
        assert result == "[Section not found: doc.md]"

    # ── block not found (F-H2-011) ──────────────────────────────────────

    def test_block_not_found(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.BLOCK_NOT_FOUND,
            "notes.md",
            block_id="abc123",
        )
        assert result == "[Block not found: notes.md#^abc123]"

    def test_block_not_found_no_block_id(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.BLOCK_NOT_FOUND,
            "notes.md",
        )
        assert result == "[Block not found: notes.md]"

    # ── file not found (F-H2-012) ───────────────────────────────────────

    def test_file_not_found(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.FILE_NOT_FOUND,
            "missing.md",
        )
        assert result == "[File not found: missing.md]"

    def test_file_not_found_with_heading(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.FILE_NOT_FOUND,
            "missing.md",
            heading="Setup",
        )
        assert result == "[File not found: missing.md#Setup]"

    # ── circular reference (F-H2-013) ───────────────────────────────────

    def test_circular_reference(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.CIRCULAR_REFERENCE,
            "a.md",
        )
        assert result == "[Circular reference: a.md]"

    def test_circular_reference_with_block_id(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.CIRCULAR_REFERENCE,
            "a.md",
            block_id="key",
        )
        assert result == "[Circular reference: a.md#^key]"

    # ── max depth (F-H2-014) ────────────────────────────────────────────

    def test_max_depth(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.MAX_DEPTH,
            "deep.md",
        )
        assert result == "[Max depth reached: deep.md]"

    # ── int kind ────────────────────────────────────────────────────────

    def test_int_kind_accepted(self) -> None:
        result = make_error_placeholder(1, "f.md", heading="H")
        assert result == "[Section not found: f.md#H]"

    # ── heading takes precedence over block_id ──────────────────────────

    def test_heading_priority(self) -> None:
        result = make_error_placeholder(
            LinkErrorKind.SECTION_NOT_FOUND,
            "f.md",
            heading="H",
            block_id="b",
        )
        assert result == "[Section not found: f.md#H]"

    # ── deterministic output ────────────────────────────────────────────

    def test_output_is_stable(self) -> None:
        a = make_error_placeholder(LinkErrorKind.FILE_NOT_FOUND, "x.md")
        b = make_error_placeholder(LinkErrorKind.FILE_NOT_FOUND, "x.md")
        assert a == b


# ═══════════════════════════════════════════════════════════════════════════
# make_keep_link  (F-H2-012 … F-H2-014 keep mode)
# ═══════════════════════════════════════════════════════════════════════════


class TestMakeKeepLink:
    """Verify wiki-embed reconstruction for keep-mode error handling."""

    def test_plain_file(self) -> None:
        assert make_keep_link("doc.md") == "![[doc.md]]"

    def test_file_with_heading(self) -> None:
        assert make_keep_link("doc.md", heading="API") == "![[doc.md#API]]"

    def test_file_with_block_id(self) -> None:
        assert make_keep_link("doc.md", block_id="abc") == "![[doc.md#^abc]]"

    def test_heading_takes_precedence_over_block_id(self) -> None:
        result = make_keep_link("doc.md", heading="H", block_id="b")
        assert result == "![[doc.md#H]]"

    def test_unicode_filename(self) -> None:
        assert make_keep_link("文档.md") == "![[文档.md]]"


# ═══════════════════════════════════════════════════════════════════════════
# dispatch_error_output  mode-aware dispatch
# ═══════════════════════════════════════════════════════════════════════════


class TestDispatchErrorOutput:
    """Verify the single-entry-point error dispatcher works correctly for
    all error kinds and all three modes."""

    # ── ignore mode ─────────────────────────────────────────────────────

    @pytest.mark.parametrize("kind", LinkErrorKind)
    def test_ignore_returns_empty_string(self, kind: LinkErrorKind) -> None:
        result = dispatch_error_output(kind, "ignore", "f.md")
        assert result == ""

    # ── keep mode ───────────────────────────────────────────────────────

    def test_keep_with_original_link(self) -> None:
        result = dispatch_error_output(
            LinkErrorKind.FILE_NOT_FOUND,
            "keep",
            "f.md",
            original_link="![[f.md|custom]]",
        )
        assert result == "![[f.md|custom]]"

    def test_keep_fallback_synthesises_link(self) -> None:
        result = dispatch_error_output(
            LinkErrorKind.CIRCULAR_REFERENCE,
            "keep",
            "loop.md",
            heading="X",
        )
        assert result == "![[loop.md#X]]"

    # ── placeholder mode ────────────────────────────────────────────────

    @pytest.mark.parametrize(
        "kind, expected_prefix",
        [
            (LinkErrorKind.SECTION_NOT_FOUND, "[Section not found:"),
            (LinkErrorKind.BLOCK_NOT_FOUND, "[Block not found:"),
            (LinkErrorKind.FILE_NOT_FOUND, "[File not found:"),
            (LinkErrorKind.CIRCULAR_REFERENCE, "[Circular reference:"),
            (LinkErrorKind.MAX_DEPTH, "[Max depth reached:"),
        ],
    )
    def test_placeholder_has_correct_prefix(
        self,
        kind: LinkErrorKind,
        expected_prefix: str,
    ) -> None:
        result = dispatch_error_output(kind, "placeholder", "f.md")
        assert result.startswith(expected_prefix)
        assert "f.md" in result

    # ── str mode accepted ───────────────────────────────────────────────

    def test_str_mode_ignore(self) -> None:
        assert (
            dispatch_error_output(
                LinkErrorKind.MAX_DEPTH,
                "ignore",
                "x.md",
            )
            == ""
        )

    def test_str_mode_keep(self) -> None:
        r = dispatch_error_output(
            LinkErrorKind.MAX_DEPTH,
            "keep",
            "x.md",
        )
        assert r == "![[x.md]]"


# ═══════════════════════════════════════════════════════════════════════════
# resolve_embedded_links — max-depth integration (F-H2-014)
# ═══════════════════════════════════════════════════════════════════════════


class TestResolveEmbeddedLinksMaxDepth:
    """End-to-end verification that ``on_max_depth`` produces correct output
    for all three modes in the batch resolver."""

    def test_default_is_placeholder(self, tmp_path: Path) -> None:
        """By default, max-depth produces a placeholder."""
        src = _write(tmp_path / "src" / "main.md", "![[a.md]]\n")
        _write(tmp_path / "src" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "src" / "b.md", "![[c.md]]\n")
        _write(tmp_path / "src" / "c.md", "Deep.\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=1,
        )
        assert "[Max depth reached:" in result
        assert "![[a.md]]" not in result

    def test_ignore_returns_empty(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", "![[a.md]]")
        _write(tmp_path / "src" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "src" / "b.md", "Deep.\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=1,
            on_max_depth="ignore",
        )
        assert "[Max depth reached:" not in result
        assert "![[a.md]]" not in result  # was expanded

    def test_keep_returns_link(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", "![[a.md]]")
        _write(tmp_path / "src" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "src" / "b.md", "Deep.\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=1,
            on_max_depth="keep",
        )
        # With max_depth=1, a.md is expanded and only its blocked child link
        # is preserved; surrounding child content must not be replaced.
        assert "![[b.md]]" in result
        assert "![[a.md]]" not in result


# ═══════════════════════════════════════════════════════════════════════════
# User-path integration: all error paths through resolve_embedded_links
# ═══════════════════════════════════════════════════════════════════════════


class TestUserPathErrorSemantics:
    """End-to-end user path tests proving error conditions produce stable,
    diagnostic output that downstream consumers can rely on."""

    # ── file not found ──────────────────────────────────────────────────

    def test_not_found_placeholder_in_mid_document(self, tmp_path: Path) -> None:
        """A document with a missing embed mid-paragraph gets a bracketed
        placeholder that does not corrupt surrounding text."""
        src = _write(tmp_path / "vault" / "doc.md", "main content\n")
        result = resolve_embedded_links(
            "Before. ![[ghost.md]] After.",
            str(src),
            on_not_found="placeholder",
        )
        assert "Before." in result
        assert "After." in result
        assert "[File not found: ghost.md]" in result
        assert "![[ghost.md]]" not in result

    def test_not_found_ignore_removes_link_quietly(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "doc.md", ".\n")
        result = resolve_embedded_links(
            "Fix: ![[ghost.md]] done.",
            str(src),
            on_not_found="ignore",
        )
        assert result.strip() == "Fix:  done."

    def test_not_found_keep_preserves_link(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "doc.md", ".\n")
        result = resolve_embedded_links(
            "See ![[ghost.md]].",
            str(src),
            on_not_found="keep",
        )
        assert "![[ghost.md]]" in result

    # ── section not found ───────────────────────────────────────────────

    def test_section_not_found_placeholder(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "main.md", "# Main\n")
        _write(tmp_path / "vault" / "lib.md", "# Lib\n\n## Existing\n\nText.\n")

        result = resolve_embedded_links(
            "![[lib.md#Missing]]",
            str(src),
            on_not_found="placeholder",
        )
        assert "[Section not found: lib.md#Missing]" in result

    # ── block not found ─────────────────────────────────────────────────

    def test_block_not_found_placeholder(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "main.md", "# Main\n")
        _write(tmp_path / "vault" / "notes.md", "# Notes\n\nText.\n")

        result = resolve_embedded_links(
            "![[notes.md#^nope]]",
            str(src),
            on_not_found="placeholder",
        )
        assert "[Block not found: notes.md#^nope]" in result

    # ── circular reference (indirect) ───────────────────────────────────

    def test_indirect_circular_placeholder(self, tmp_path: Path) -> None:
        """A → B → A produces a circular-reference placeholder in B's expansion."""
        src = _write(tmp_path / "vault" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "vault" / "b.md", "![[a.md]]\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            on_circular="placeholder",
        )
        assert "[Circular reference:" in result

    def test_circular_keep(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "vault" / "b.md", "![[a.md]]\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            on_circular="keep",
        )
        assert "![[a.md]]" in result

    def test_circular_ignore(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "vault" / "b.md", "![[a.md]]\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            on_circular="ignore",
        )
        assert "[Circular reference:" not in result
        assert "![[a.md]]" not in result

    # ── max depth ───────────────────────────────────────────────────────

    def test_max_depth_default_placeholder(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "main.md", "![[a.md]]\n")
        _write(tmp_path / "vault" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "vault" / "b.md", "![[c.md]]\n")
        _write(tmp_path / "vault" / "c.md", "Deep.\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=1,
        )
        assert "[Max depth reached:" in result

    def test_max_depth_user_can_configure_ignore(self, tmp_path: Path) -> None:
        """User passing on_max_depth='ignore' gets silent truncation with
        no placeholder leak into the output."""
        src = _write(tmp_path / "vault" / "main.md", "![[a.md]]\n")
        _write(tmp_path / "vault" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "vault" / "b.md", "![[c.md]]\n")
        _write(tmp_path / "vault" / "c.md", "Deep.\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=1,
            on_max_depth="ignore",
        )
        assert "[Max depth reached:" not in result


# ═══════════════════════════════════════════════════════════════════════════
# NotFoundAction enum
# ═══════════════════════════════════════════════════════════════════════════


class TestNotFoundActionEnum:
    """Verify the enum has exactly three actions."""

    def test_all_actions_present(self) -> None:
        assert set(NotFoundAction) == {"ignore", "keep", "placeholder"}

    def test_values_match_strings(self) -> None:
        assert NotFoundAction.IGNORE.value == "ignore"
        assert NotFoundAction.KEEP.value == "keep"
        assert NotFoundAction.PLACEHOLDER.value == "placeholder"

    def test_str_mode_construction(self) -> None:
        assert NotFoundAction("ignore") == NotFoundAction.IGNORE
        assert NotFoundAction("keep") == NotFoundAction.KEEP
        assert NotFoundAction("placeholder") == NotFoundAction.PLACEHOLDER

    def test_invalid_mode_raises(self) -> None:
        with pytest.raises(ValueError):
            NotFoundAction("bogus")
