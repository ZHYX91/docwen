"""Focused tests split from test_links_embed_markdown.py."""

from __future__ import annotations

from ._links_embed_markdown_support import (
    Path,
    _write,
    process_embedded_md_file,
    pytest,
)

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    ("line", "expected"),
    [
        ("| header | value |", True),
        ("> | quoted | row |", True),
        (r"name | escaped \| value |", True),
        ("| --- | :---: |", False),
        (r"\| escaped delimiters only \|", False),
        ("plain prose", False),
    ],
)
def test_probable_table_row_requires_real_unescaped_cells(line: str, expected: bool) -> None:
    """F-H2-027: table-safe embed routing distinguishes data rows from separators and prose."""
    from docwen_core.links._embed_dispatch import _is_probable_table_row_line

    assert _is_probable_table_row_line(line) is expected


class TestProcessEmbeddedMdFile:
    """Covers F-H2-009: ``process_embedded_md_file``."""

    # ── basic embed (whole file) ──────────────────────────────────────

    def test_embed_whole_file(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# Main\n\nHello world.\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
        )
        assert "Hello world." in result
        assert "# Main" in result

    def test_embed_strips_yaml_front_matter(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", ("---\ntitle: Doc\n---\n\n# Heading\nBody text.\n"))
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
        )
        assert "title:" not in result
        assert "# Heading" in result
        assert "Body text." in result

    # ── section extraction ────────────────────────────────────────────

    def test_embed_heading_section(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", ("# Top\n\nIntro.\n\n## Features\n\nFeature list.\n\n## Limits\n\n"))
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            heading="Features",
        )
        assert "## Features" in result
        assert "Feature list." in result
        assert "## Limits" not in result
        assert "# Top" not in result

    def test_embed_heading_not_found_placeholder(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# Top\n\nBody.\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            heading="Missing",
            on_not_found="placeholder",
        )
        assert "Section not found" in result
        assert "main.md#Missing" in result

    def test_embed_heading_not_found_ignore(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# Top\n\nBody.\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            heading="Missing",
            on_not_found="ignore",
        )
        assert result == ""

    def test_embed_heading_not_found_keep(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# Top\n\nBody.\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            heading="Missing",
            on_not_found="keep",
        )
        assert result == "![[main.md#Missing]]"

    # ── block extraction ──────────────────────────────────────────────

    def test_embed_block_id(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", ("# Notes\n\nKey insight here. ^key1\n\nOther text.\n"))
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            block_id="key1",
        )
        assert "Key insight here." in result
        assert "^key1" not in result
        assert "Other text." not in result

    def test_embed_block_id_standalone(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", ("# Notes\n\nThe block contents.\n\n^b1\n\nAfter.\n"))
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            block_id="b1",
        )
        assert "The block contents." in result
        assert "^b1" not in result

    def test_embed_block_not_found_placeholder(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# Notes\n\nText.\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            block_id="nope",
            on_not_found="placeholder",
        )
        assert "Block not found" in result
        assert "main.md#^nope" in result

    # ── mode: keep ────────────────────────────────────────────────────

    def test_mode_keep(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# N\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="keep",
        )
        assert result == "![[main.md]]"

    def test_mode_keep_with_heading(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# N\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="keep",
            heading="Intro",
        )
        assert result == "![[main.md#Intro]]"

    def test_mode_keep_with_block_id(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# N\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="keep",
            block_id="abc",
        )
        assert result == "![[main.md#^abc]]"

    # ── mode: extract_text ────────────────────────────────────────────

    def test_mode_extract_text_with_display(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# N\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="extract_text",
            display_text="My Title",
        )
        assert result == "My Title"

    def test_mode_extract_text_fallback_to_stem(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# N\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="extract_text",
        )
        assert result == "main"

    # ── mode: remove ──────────────────────────────────────────────────

    def test_mode_remove(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "src" / "main.md", "# N\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="remove",
        )
        assert result == ""

    # ── file not found ────────────────────────────────────────────────

    def test_file_not_found_placeholder(self) -> None:
        result = process_embedded_md_file(
            "/nonexistent/file.md",
            "/src/main.md",
            set(),
            0,
            mode="embed",
            on_not_found="placeholder",
        )
        assert "File not found" in result

    def test_file_not_found_ignore(self) -> None:
        result = process_embedded_md_file(
            "/nonexistent/file.md",
            "/src/main.md",
            set(),
            0,
            mode="embed",
            on_not_found="ignore",
        )
        assert result == ""

    def test_file_not_found_keep(self) -> None:
        result = process_embedded_md_file(
            "/nonexistent/file.md",
            "/src/main.md",
            set(),
            0,
            mode="embed",
            on_not_found="keep",
        )
        assert result == "![[file.md]]"

    # ── circular reference detection ──────────────────────────────────

    def test_circular_detected_placeholder(self, tmp_path: Path) -> None:
        """A embeds A — direct self-reference detected via visited_files."""
        md = _write(tmp_path / "a.md", "# A\n")
        resolved = str(Path(md).resolve())
        visited: set[str] = {resolved}
        result = process_embedded_md_file(
            md,
            md,
            visited,
            0,
            mode="embed",
            on_circular="placeholder",
        )
        assert "Circular reference" in result

    def test_circular_detected_ignore(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "a.md", "# A\n")
        visited: set[str] = {str(Path(md).resolve())}
        result = process_embedded_md_file(
            md,
            md,
            visited,
            0,
            mode="embed",
            on_circular="ignore",
        )
        assert result == ""

    def test_circular_detected_keep(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "a.md", "# A\n")
        visited: set[str] = {str(Path(md).resolve())}
        result = process_embedded_md_file(
            md,
            md,
            visited,
            0,
            mode="embed",
            on_circular="keep",
        )
        assert result == "![[a.md]]"

    def test_circular_disabled(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "a.md", "# A\n")
        visited: set[str] = {str(Path(md).resolve())}
        result = process_embedded_md_file(
            md,
            md,
            visited,
            0,
            mode="embed",
            detect_circular=False,
        )
        # No circular-detection → reads the file normally
        assert "# A" in result

    # ── recursive expansion via callback ──────────────────────────────

    def test_recursive_expansion_with_callback(self, tmp_path: Path) -> None:
        """The process_links callback is invoked for nested embeds."""
        md = _write(tmp_path / "inner.md", "# Inner\n\nNested content.\n")

        calls: list[str] = []

        def _fake_scanner(
            content: str,
            source_path: str,
            visited: set[str] | None,
            depth_: int,
            **kwargs: object,
        ) -> str:
            calls.append(content)
            return content.upper()

        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            process_links=_fake_scanner,
        )
        assert len(calls) == 1
        assert "Nested content." in calls[0]
        assert result == calls[0].upper()

    def test_no_callback_returns_raw(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "inner.md", "# Inner\n\nContent.\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
        )
        assert "Content." in result
        assert "# Inner" in result

    # ── mode strings accepted ─────────────────────────────────────────

    def test_str_mode_accepted(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "x.md", "# X\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
        )
        assert "# X" in result

    def test_unknown_mode_falls_back_to_embed(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "x.md", "# X\n")
        with pytest.raises(ValueError):
            process_embedded_md_file(md, md, set(), 0, mode="bogus")

    def test_not_found_str_action_accepted(self, tmp_path: Path) -> None:
        md = _write(tmp_path / "x.md", "# X\n")
        result = process_embedded_md_file(
            md,
            md,
            set(),
            0,
            mode="embed",
            heading="Nope",
            on_not_found="ignore",
        )
        assert result == ""
