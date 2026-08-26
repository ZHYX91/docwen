"""Focused tests split from test_links_embed_markdown.py."""

from __future__ import annotations

import pytest

from ._links_embed_markdown_support import (
    _SAMPLE_PNG_BYTES,
    Path,
    _write,
    process_single_embed,
    resolve_embedded_links,
)

pytestmark = pytest.mark.unit


class TestUserPathEmbedMdToDocx:
    """End-to-end user path: Markdown file with embeds → resolved output."""

    def test_full_embed_pipeline(self, tmp_path: Path) -> None:
        """Simulate the real user path:
        1. Main document has ![[child.md]]
        2. resolve_embedded_links expands it
        3. The expanded content includes child's content
        """
        src = _write(tmp_path / "vault" / "report.md", "# Report\n\n![[section.md]]\n")
        _write(
            tmp_path / "vault" / "section.md",
            ("---\ntitle: Section\n---\n\n## Analysis\n\nKey findings here.\n\n## Notes\n\nSome notes.\n"),
        )

        resolved = resolve_embedded_links(
            Path(src).read_text(encoding="utf-8"),
            src,
        )
        assert "## Analysis" in resolved
        assert "Key findings here." in resolved
        assert "title:" not in resolved
        assert "![[section.md]]" not in resolved

    def test_precision_embed_pipeline(self, tmp_path: Path) -> None:
        """User embeds only a specific section via ![[file.md#heading]]."""
        src = _write(tmp_path / "vault" / "main.md", "# Main\n\n![[lib.md#API]]\n")
        _write(
            tmp_path / "vault" / "lib.md",
            ("# Library\n\nPublic.\n\n## API\n\nAPI reference.\n\n## Internal\n\nHidden.\n"),
        )

        resolved = resolve_embedded_links(
            Path(src).read_text(encoding="utf-8"),
            src,
        )
        assert "API reference." in resolved
        assert "Hidden." not in resolved
        assert "## Internal" not in resolved

    def test_block_embed_pipeline(self, tmp_path: Path) -> None:
        """User embeds a specific block via ![[file.md#^block-id]]."""
        src = _write(tmp_path / "vault" / "main.md", "# Main\n\n![[notes.md#^key]]\n")
        _write(
            tmp_path / "vault" / "notes.md", ("# Notes\n\nSome text.\n\nCritical insight here. ^key\n\nOther notes.\n")
        )

        resolved = resolve_embedded_links(
            Path(src).read_text(encoding="utf-8"),
            src,
        )
        assert "Critical insight here." in resolved
        assert "Other notes." not in resolved

    def test_multi_level_embed_pipeline(self, tmp_path: Path) -> None:
        """Three-level embed: A → B → C."""
        src = _write(tmp_path / "vault" / "a.md", "# A\n\n![[b.md]]\n")
        _write(tmp_path / "vault" / "b.md", "# B\n\n![[c.md#Detail]]\n")
        _write(tmp_path / "vault" / "c.md", ("# C\n\nIntro.\n\n## Detail\n\nDeep detail.\n\n## Extra\n\nExtra.\n"))

        resolved = resolve_embedded_links(
            Path(src).read_text(encoding="utf-8"),
            src,
        )
        assert "# A" in resolved
        assert "# B" in resolved
        assert "Deep detail." in resolved
        assert "Extra." not in resolved


class TestUserPathSingleEmbedDispatch:
    """Verify that ``process_single_embed`` correctly routes different
    embed types through a single entry point."""

    def test_dispatch_image_creates_placeholder(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "doc.md", "# Doc\n")
        img = tmp_path / "vault" / "logo.png"
        img.parent.mkdir(parents=True, exist_ok=True)
        img.write_bytes(_SAMPLE_PNG_BYTES)

        result = process_single_embed(
            "logo.png",
            "![[logo.png]]",
            str(src),
            set(),
            0,
            image_mode="embed",
        )
        assert result is not None
        # The resolved path is absolute; the placeholder carries the full path.
        assert result.startswith("{{IMAGE:")
        assert "logo.png" in result

    def test_dispatch_md_heading_returns_section(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "doc.md", "# Doc\n")
        _write(tmp_path / "vault" / "ref.md", ("# Ref\n\nTop.\n\n## Target\n\nTarget content.\n"))
        result = process_single_embed(
            "ref.md#Target",
            "![[ref.md#Target]]",
            str(src),
            set(),
            0,
            md_mode="embed",
        )
        assert result is not None
        assert "Target content." in result
        assert "Top." not in result

    def test_dispatch_unknown_returns_none(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "vault" / "doc.md", "# Doc\n")
        unknown = tmp_path / "vault" / "data.xyz"
        unknown.parent.mkdir(parents=True, exist_ok=True)
        unknown.write_bytes(b"\x00\x01\x02\x03" * 8)

        result = process_single_embed(
            "data.xyz",
            "![[data.xyz]]",
            str(src),
            set(),
            0,
        )
        assert result is None
