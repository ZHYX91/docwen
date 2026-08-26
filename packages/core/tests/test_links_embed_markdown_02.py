"""Focused tests split from test_links_embed_markdown.py."""

from __future__ import annotations

from ._links_embed_markdown_support import (
    _SAMPLE_PNG_BYTES,
    EmbeddedMdMode,
    NotFoundAction,
    Path,
    _data_uri,
    _sample_image_bytes,
    _write,
    process_single_embed,
    pytest,
    resolve_embedded_links,
)

pytestmark = pytest.mark.unit


class TestProcessSingleEmbed:
    """Covers F-H2-026: ``process_single_embed`` dispatch."""

    # ── data URI image dispatch ───────────────────────────────────────

    def test_data_uri_image_dispatches_to_image_processor(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        result = process_single_embed(
            _data_uri(),
            "![[data-uri]]",
            src,
            set(),
            0,
            image_mode="keep",
        )
        # keep mode → the original link text is preserved
        assert result == "![[data-uri]]"

    def test_data_uri_image_embed_mode(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        result = process_single_embed(
            _data_uri(),
            "![[data-uri.png]]",
            src,
            set(),
            0,
            image_mode="embed",
        )
        assert result is not None
        assert result.startswith("{{IMAGE:")

    def test_data_uri_decode_failure_extract_text(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        result = process_single_embed(
            "data:image/png;base64,!!!bad!!!",
            "![[bad]]",
            src,
            set(),
            0,
            image_mode="extract_text",
            display_text="Alt",
        )
        assert result == "Alt"

    # ── markdown file dispatch ────────────────────────────────────────

    def test_dispatches_md_file_embed(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n![[sub.md]]\n")
        _write(tmp_path / "src" / "sub.md", "# Sub\n\nSub content.\n")
        result = process_single_embed(
            "sub.md",
            "![[sub.md]]",
            src,
            set(),
            0,
            md_mode="embed",
        )
        assert result is not None
        assert "Sub content." in result

    def test_dispatches_md_file_with_heading(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        _write(tmp_path / "src" / "sub.md", "# Top\n\n## Details\n\nInfo.\n")
        result = process_single_embed(
            "sub.md#Details",
            "![[sub.md#Details]]",
            src,
            set(),
            0,
            md_mode="embed",
        )
        assert result is not None
        assert "## Details" in result
        assert "# Top" not in result

    def test_dispatches_md_file_with_block_id(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        _write(tmp_path / "src" / "sub.md", "# Sub\n\nKey point. ^kp\n")
        result = process_single_embed(
            "sub.md#^kp",
            "![[sub.md#^kp]]",
            src,
            set(),
            0,
            md_mode="embed",
        )
        assert result is not None
        assert "Key point." in result
        assert "^kp" not in result

    # ── image file dispatch ───────────────────────────────────────────

    def test_dispatches_image_file(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        img = tmp_path / "src" / "photo.png"
        img.parent.mkdir(parents=True, exist_ok=True)
        img.write_bytes(_SAMPLE_PNG_BYTES)

        result = process_single_embed(
            "photo.png",
            "![[photo.png]]",
            str(src),
            set(),
            0,
            image_mode="keep",
        )
        assert result == "![[photo.png]]"

    def test_dispatches_image_file_with_dimensions(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        img = tmp_path / "src" / "pic.png"
        img.parent.mkdir(parents=True, exist_ok=True)
        img.write_bytes(_SAMPLE_PNG_BYTES)

        result = process_single_embed(
            "pic.png",
            "![[pic.png|300x200]]",
            str(src),
            set(),
            0,
            image_mode="embed",
            width=300,
            height=200,
        )
        assert result is not None
        assert "{{IMAGE:" in result
        assert "|300|200" in result

    @pytest.mark.parametrize(
        "image_format",
        ["JPEG", "PNG", "GIF", "BMP", "TIFF", "WEBP"],
    )
    def test_dispatches_image_content_with_unrecognized_suffix(
        self,
        tmp_path: Path,
        image_format: str,
    ) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        image = tmp_path / "src" / f"{image_format.lower()}.resource"
        image.write_bytes(_sample_image_bytes(image_format))

        result = process_single_embed(
            image.name,
            f"![[{image.name}]]",
            str(src),
            set(),
            0,
            image_mode="embed",
        )

        assert result is not None
        assert result.startswith("{{IMAGE:")
        assert str(image) in result

    @pytest.mark.parametrize(
        "content",
        ["# Embedded\n\nContent-first routing.\n", "Plain embedded text.\n"],
        ids=["markdown", "plain-text"],
    )
    def test_dispatches_text_content_with_unrecognized_suffix(
        self,
        tmp_path: Path,
        content: str,
    ) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        embedded = tmp_path / "src" / "section.resource"
        embedded.write_text(content, encoding="utf-8")

        result = process_single_embed(
            embedded.name,
            f"![[{embedded.name}]]",
            str(src),
            set(),
            0,
            md_mode="embed",
        )

        assert result is not None
        assert content.strip() in result

    def test_image_suffix_does_not_admit_corrupt_image_content(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        disguised = tmp_path / "src" / "not-an-image.png"
        disguised.write_bytes(b"\x89PNG\r\n\x1a\ntruncated-image")

        result = process_single_embed(
            disguised.name,
            f"![[{disguised.name}]]",
            str(src),
            set(),
            0,
        )

        assert result is None

    # ── unknown file type ─────────────────────────────────────────────

    def test_unknown_file_type_returns_none(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        unknown = tmp_path / "src" / "data.bin"
        unknown.parent.mkdir(parents=True, exist_ok=True)
        unknown.write_bytes(b"\x00\x01\x02\x03" * 8)

        result = process_single_embed(
            "data.bin",
            "![[data.bin]]",
            str(src),
            set(),
            0,
        )
        assert result is None

    # ── file not found ────────────────────────────────────────────────

    def test_file_not_found_placeholder(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        result = process_single_embed(
            "ghost.md",
            "![[ghost.md]]",
            src,
            set(),
            0,
            on_not_found="placeholder",
        )
        assert result is not None
        assert "File not found" in result

    def test_file_not_found_ignore(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        result = process_single_embed(
            "ghost.md",
            "![[ghost.md]]",
            src,
            set(),
            0,
            on_not_found="ignore",
        )
        assert result == ""

    def test_file_not_found_keep(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "note.md", "# Note\n")
        result = process_single_embed(
            "ghost.md",
            "![[ghost.md]]",
            src,
            set(),
            0,
            on_not_found="keep",
        )
        assert result == "![[ghost.md]]"


class TestResolveEmbeddedLinks:
    """Batch resolution of ``![[...]]`` links in Markdown content."""

    # ── basic resolution ──────────────────────────────────────────────

    def test_resolves_single_md_embed(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", "Content.\n")
        _write(tmp_path / "src" / "child.md", "# Child\n\nHi from child.\n")

        result = resolve_embedded_links(
            "Before.\n![[child.md]]\nAfter.",
            str(src),
        )
        assert "Hi from child." in result
        assert "![[child.md]]" not in result

    def test_resolves_multiple_embeds(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        _write(tmp_path / "src" / "a.md", "Content A.\n")
        _write(tmp_path / "src" / "b.md", "Content B.\n")

        result = resolve_embedded_links(
            "![[a.md]] and ![[b.md]]",
            str(src),
        )
        assert "Content A." in result
        assert "Content B." in result
        assert "![[a.md]]" not in result
        assert "![[b.md]]" not in result

    # ── heading / block-id in batch ───────────────────────────────────

    def test_resolves_embed_with_heading(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        _write(tmp_path / "src" / "doc.md", ("# Top\n\nIntro.\n\n## API\n\nAPI docs.\n\n## FAQ\n\nQ&A.\n"))
        result = resolve_embedded_links(
            "![[doc.md#API]]",
            str(src),
        )
        assert "API docs." in result
        assert "FAQ" not in result

    def test_resolves_embed_with_block_id(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        _write(tmp_path / "src" / "doc.md", ("# Notes\n\nImportant note. ^imp\n\nOther.\n"))
        result = resolve_embedded_links(
            "![[doc.md#^imp]]",
            str(src),
        )
        assert "Important note." in result
        assert "Other." not in result

    # ── display text parsing ──────────────────────────────────────────

    def test_embed_with_display_text(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        _write(tmp_path / "src" / "child.md", "Child content.\n")
        # With md_mode="extract_text", display_text is used
        result = resolve_embedded_links(
            "![[child.md|My Label]]",
            str(src),
            md_mode="extract_text",
        )
        assert result == "My Label"

    def test_embed_with_size_display(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        img = tmp_path / "src" / "pic.png"
        img.parent.mkdir(parents=True, exist_ok=True)
        img.write_bytes(_SAMPLE_PNG_BYTES)

        result = resolve_embedded_links(
            "![[pic.png|200x150]]",
            str(src),
            image_mode="embed",
        )
        assert "{{IMAGE:" in result
        assert "|200|150" in result

    # ── recursive (nested) embeds ─────────────────────────────────────

    def test_recursive_nested_embed(self, tmp_path: Path) -> None:
        """A embeds B, and B embeds C — all should be expanded."""
        src = _write(tmp_path / "src" / "main.md", "Main.\n")
        _write(tmp_path / "src" / "b.md", "B before.\n![[c.md]]\nB after.\n")
        _write(tmp_path / "src" / "c.md", "C content.\n")

        result = resolve_embedded_links(
            "![[b.md]]",
            str(src),
        )
        assert "B before." in result
        assert "C content." in result
        assert "B after." in result
        assert "![[c.md]]" not in result

    def test_recursive_heading_embed(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", "Main.\n")
        _write(tmp_path / "src" / "outer.md", ("# Outer\n\n![[inner.md#Detail]]\n"))
        _write(tmp_path / "src" / "inner.md", ("# Inner\n\nTop.\n\n## Detail\n\nDetail text.\n\n## More\n\nMore.\n"))
        result = resolve_embedded_links(
            "![[outer.md]]",
            str(src),
        )
        assert "Detail text." in result
        assert "More." not in result
        assert "![[inner.md#Detail]]" not in result

    # ── max depth ─────────────────────────────────────────────────────

    def test_max_depth_honoured_placeholder(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", "![[a.md]]\n")
        _write(tmp_path / "src" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "src" / "b.md", "![[c.md]]\n")
        _write(tmp_path / "src" / "c.md", "Deep.\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=2,
            on_max_depth="placeholder",
        )
        assert "[Max depth reached: c.md]" in result
        assert "![[b.md]]" not in result

    def test_max_depth_honoured_ignore(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", "![[a.md]]\n")
        _write(tmp_path / "src" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "src" / "b.md", "![[c.md]]\n")
        _write(tmp_path / "src" / "c.md", "Deep.\n")

        # on_max_depth="ignore" preserves the old silent behavior:
        # the content passed to the recursive call is returned as-is.
        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=2,
            on_max_depth="ignore",
        )
        assert "![[c.md]]" not in result
        assert "[Max depth reached" not in result
        # Sanity: "![[a.md]]" is expanded through to b.md's content,
        # minus c.md (stopped by max depth).
        assert "![[a.md]]" not in result

    def test_max_depth_honoured_keep(self, tmp_path: Path) -> None:
        """max_depth with keep mode returns the wiki link syntax."""
        src = _write(tmp_path / "src" / "main.md", "![[a.md]]\n")
        _write(tmp_path / "src" / "a.md", "![[b.md]]\n")
        _write(tmp_path / "src" / "b.md", "![[c.md]]\n")
        _write(tmp_path / "src" / "c.md", "Deep.\n")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=2,
            on_max_depth="keep",
        )
        assert "![[c.md]]" in result

    # ── missing file in batch ─────────────────────────────────────────

    def test_missing_file_placeholder(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        result = resolve_embedded_links(
            "![[ghost.md]]",
            str(src),
            on_not_found="placeholder",
        )
        assert "File not found" in result

    def test_missing_file_ignore(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        result = resolve_embedded_links(
            "![[ghost.md]]",
            str(src),
            on_not_found="ignore",
        )
        assert result == ""

    # ── no embeds ─────────────────────────────────────────────────────

    def test_content_without_embeds_unchanged(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        text = "Just plain markdown.\n\n# Heading\n\nNo embeds here.\n"
        result = resolve_embedded_links(text, str(src))
        assert result == text

    def test_empty_content(self, tmp_path: Path) -> None:
        result = resolve_embedded_links("", "/tmp/fake.md")
        assert result == ""

    # ── data URI in batch ─────────────────────────────────────────────

    def test_data_uri_in_batch_resolution(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "src" / "main.md", ".\n")
        result = resolve_embedded_links(
            f"![[{_data_uri()}]]",
            str(src),
            image_mode="keep",
        )
        assert "data:image/png;base64," in result


class TestEmbeddedMdMode:
    def test_all_modes_present(self) -> None:
        assert set(EmbeddedMdMode) == {"embed", "keep", "extract_text", "remove"}

    def test_values_match_strings(self) -> None:
        assert EmbeddedMdMode.EMBED.value == "embed"
        assert EmbeddedMdMode.KEEP.value == "keep"
        assert EmbeddedMdMode.EXTRACT_TEXT.value == "extract_text"
        assert EmbeddedMdMode.REMOVE.value == "remove"


class TestNotFoundAction:
    def test_all_actions_present(self) -> None:
        assert set(NotFoundAction) == {"ignore", "keep", "placeholder"}

    def test_values_match_strings(self) -> None:
        assert NotFoundAction.IGNORE.value == "ignore"
        assert NotFoundAction.KEEP.value == "keep"
        assert NotFoundAction.PLACEHOLDER.value == "placeholder"
