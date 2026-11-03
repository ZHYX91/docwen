"""Markdown consumers recognize the shared Number Suite provenance format."""

from types import SimpleNamespace

import pytest

from docwen_plugin_markdown.common_utils import add_md_numbering, remove_md_numbering

pytestmark = pytest.mark.contract
WJ = "\u2060"


class _ExactSchemeRegistry:
    def get_scheme(self, scheme_id):
        assert scheme_id == "exact"
        return SimpleNamespace(enabled=True, levels={"level_1": "{1.arabic_half} "})


def test_markdown_cleanup_recognizes_only_heading_markers():
    source = f"# {WJ}第一章{WJ} Title\n\nText {WJ}1.{WJ} stays\n\n```md\n# {WJ}1.{WJ} code\n```\n"
    assert remove_md_numbering(source) == f"# Title\n\nText {WJ}1.{WJ} stays\n\n```md\n# {WJ}1.{WJ} code\n```\n"


def test_markdown_add_only_preserves_marked_and_incomplete_headings():
    source = f"# {WJ}9.{WJ} Title\n\n# {WJ}broken\n\n# Next\n"
    assert (
        add_md_numbering(source, "exact", _ExactSchemeRegistry()) == f"# {WJ}9.{WJ} Title\n\n# {WJ}broken\n\n# 3 Next\n"
    )
