"""Exact public lexical ownership without changing ordinary link rewriting."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import (
    is_markdown_source_escaped,
    markdown_inline_source_owners,
    markdown_source_owners,
    process_markdown_links,
    split_markdown_inline_segments,
)

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "link",
    [
        "[中文😀](local(foo)%%draft)",
        r"[中文😀](local\(foo\)%%draft)",
        '[中文😀](local(foo) "标题 ) %%draft")',
        "[中文😀](<local(foo)%%draft>)",
        "![中文😀](local(foo)%%draft)",
    ],
)
def test_parsed_metadata_owns_the_complete_balanced_target(link: str) -> None:
    source = "前文\r\n" + link + "\r\n后文"
    [owner] = markdown_inline_source_owners(source)
    assert owner.kind == "link_metadata"
    assert source[owner.start : owner.end] == link[link.index("](") :]
    assert split_markdown_inline_segments(source, protect_bare_urls=False) == [(source, False)]


def test_orphan_destination_does_not_gain_parsed_link_ownership() -> None:
    assert markdown_inline_source_owners("](local(foo)%%draft)") == []


def test_label_atoms_and_metadata_have_original_source_offsets() -> None:
    source = "前文😀\r\n[示例 `%%` $a<!--b$ visible](local(foo)%%draft) 后文"
    owners = markdown_inline_source_owners(source)
    assert [(source[owner.start : owner.end], owner.kind) for owner in owners] == [
        ("`%%`", "literal"),
        ("$a<!--b$", "literal"),
        ("](local(foo)%%draft)", "link_metadata"),
    ]
    assert owners[0].start < owners[1].start < owners[2].start


@pytest.mark.parametrize("wiki", ["[[Chapter%%Draft]]", "[[Chapter|Alias<!--Draft]]", "![[Chapter%%Draft]]"])
def test_wiki_ownership_is_distinct_from_renderer_literal_shielding(wiki: str) -> None:
    [owner] = markdown_inline_source_owners(wiki)
    assert (owner.start, owner.end, owner.kind) == (0, len(wiki), "wikilink")
    assert split_markdown_inline_segments(wiki) == [(wiki, False)]


def test_angle_prose_is_visible_but_html_and_autolinks_have_precise_owners() -> None:
    assert markdown_inline_source_owners("当 x < 2 @visible > 0") == []
    for source in ['<span data-x="%%">', "<https://example.test/%%draft>"]:
        [owner] = markdown_inline_source_owners(source)
        assert (owner.start, owner.end, owner.kind) == (0, len(source), "literal")


@pytest.mark.parametrize("slashes", [1, 2, 3, 4])
@pytest.mark.parametrize("html", ["<!-- literal -->", '<span data-x="%%">'])
def test_html_renderer_atoms_follow_source_escape_parity(slashes: int, html: str) -> None:
    source = "\\" * slashes + html
    assert is_markdown_source_escaped(source, slashes) == bool(slashes % 2)
    owners = markdown_inline_source_owners(source)
    if slashes % 2:
        assert owners == []
        assert split_markdown_inline_segments(source) == [(source, False)]
    else:
        [owner] = owners
        assert source[owner.start : owner.end] == html


@pytest.mark.parametrize("source", ["prefixhttps://example.test/%%", "https://%%"])
def test_invalid_bare_urls_have_no_lexical_owner(source: str) -> None:
    assert markdown_inline_source_owners(source) == []


def test_bare_url_owner_uses_the_existing_validated_boundary() -> None:
    source = "Text https://example.test/%%draft, next"
    [owner] = markdown_inline_source_owners(source)
    assert owner.kind == "url"
    assert source[owner.start : owner.end] == "https://example.test/%%draft"


def test_comment_ownership_retains_the_original_metadata_closer() -> None:
    source = "%% [[Other%%]] [[Chapter%%Draft]]"
    owners = markdown_source_owners(source)
    assert [(source[owner.start : owner.end], owner.kind) for owner in owners] == [
        ("%% [[Other%%", "comment"),
        ("[[Chapter%%Draft]]", "wikilink"),
    ]


@pytest.mark.parametrize("mode", ["extract_text", "hyperlink"])
def test_generated_wiki_metadata_stays_literal_before_downstream_note_scans(mode: str, tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("body", encoding="utf-8")
    result = process_markdown_links(
        "[[Chapter|Alias%%Draft[^fake]<!--]]",
        str(source),
        target_format="docx",
        link_config=replace(LinkRuntimeConfig(), non_embed_wiki_mode=mode),
        declared_wiki_link=lambda _raw, _target: "obsidian://open?vault=test&file=Chapter",
        protect_source_comments=True,
    )
    assert "%%" not in result
    assert "[^fake]" not in result
    assert not any(owner.kind == "comment" for owner in markdown_source_owners(result))


@pytest.mark.parametrize("literal", ["%% [[Other%%]]", "%% [示例](local(foo)%%draft)"])
def test_source_comment_protection_preserves_original_bytes_across_link_presentation(
    literal: str, tmp_path: Path
) -> None:
    result = process_markdown_links(
        literal,
        str(tmp_path / "source.md"),
        target_format="docx",
        link_config=LinkRuntimeConfig(),
        protect_source_comments=True,
    )
    assert result == literal


@pytest.mark.parametrize("prefix", ["", "> ", "- "])
@pytest.mark.parametrize("gap", ["\n", "\n\n"])
def test_source_comment_does_not_create_html_block_authority(prefix: str, gap: str, tmp_path: Path) -> None:
    source = f"{prefix}%% hidden %%{gap}{prefix}[Visible](https://example.test/page)"
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        target_format="docx",
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="extract_text"),
        protect_source_comments=True,
    )
    assert result == f"{prefix}%% hidden %%{gap}{prefix}Visible"


@pytest.mark.parametrize("target", ["local(foo)%%draft", "local[^fake]"])
def test_empty_markdown_label_extracts_literal_destination(target: str, tmp_path: Path) -> None:
    result = process_markdown_links(
        f"[]({target})",
        str(tmp_path / "source.md"),
        target_format="docx",
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="extract_text"),
        protect_source_comments=True,
    )
    assert "%%" not in result
    assert "[^fake]" not in result
    assert "<!--" not in result
