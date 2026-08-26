"""Focused tests split from test_link_runtime_processing_contract.py."""

from __future__ import annotations

from ._link_runtime_processing_contract_support import (
    LinkRuntimeConfig,
    Path,
    _write,
    _write_png,
    escape_markdown_source_literal,
    process_markdown_links,
    pytest,
    replace,
    resolve_embedded_links,
)

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("label", ["outer [inner]", r"outer \] bracket"])
def test_bare_url_policy_does_not_rewrite_an_explicit_link_destination(
    tmp_path: Path,
    label: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"before [{label}](https://example.com) after"
    config = replace(
        LinkRuntimeConfig(),
        non_embed_markdown_mode="hyperlink",
        auto_link_bare_url=True,
    )

    result = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == ("before " + f"[{label}](<https://example.com>)" + " after")


@pytest.mark.parametrize("alt", ["outer [inner]", r"outer \] bracket"])
def test_markdown_image_policy_handles_nested_or_escaped_alt(
    tmp_path: Path,
    alt: str,
) -> None:
    source = _write(tmp_path / "source.md")
    _write_png(tmp_path / "image.png")
    text = f"before ![{alt}](image.png) after"
    config = replace(
        LinkRuntimeConfig(),
        embed_markdown_image_mode="remove",
    )

    result = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == "before  after"


def test_sibling_embed_tokens_do_not_rewrite_literal_child_text(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    literal = "Prefix text A:\x00DOCWEN_EMBED_2\x00 suffix text"
    _write(tmp_path / "a.md", literal)
    _write(tmp_path / "b.md", "B")
    _write(tmp_path / "c.md", "C")
    text = "![[a.md]] / ![[b.md]] / ![[c.md]]"

    result = process_markdown_links(
        text,
        source,
        link_config=LinkRuntimeConfig(),
    )

    assert result == f"{literal} / B / C"


def test_public_embed_resolver_rebuilds_identical_siblings_by_position(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "a.md", "`![[a.md]]`")

    result = resolve_embedded_links(
        "![[a.md]] and ![[a.md]]",
        source,
    )

    assert result == "`![[a.md]]` and `![[a.md]]`"


def test_embed_boundary_cannot_manufacture_a_policy_bypassing_link(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "child.md", "[x](")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(
        "![[child.md]]https://evil.example)",
        source,
        link_config=config,
        target_format="docx",
    )

    assert "[x](https://evil.example)" not in result


def test_embed_boundary_cannot_manufacture_a_local_image_capability(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "child.md", "[alt](")
    _write_png(tmp_path / "pixel.png")
    config = replace(LinkRuntimeConfig(), embed_markdown_image_mode="remove")

    result = process_markdown_links(
        "!![[child.md]]pixel.png)",
        source,
        link_config=config,
        target_format="docx",
    )

    assert "![alt](pixel.png)" not in result


@pytest.mark.parametrize(
    ("wiki_mode", "expected"),
    [
        ("keep", "literal"),
        ("extract_text", "text"),
        ("remove", "removed"),
        ("hyperlink", "hyperlink"),
    ],
)
def test_embed_boundary_does_not_reprocess_complete_child_wiki_output(
    tmp_path: Path,
    wiki_mode: str,
    expected: str,
) -> None:
    source = _write(tmp_path / "source.md")
    target = _write(tmp_path / "target.md", "target")
    _write(tmp_path / "child.md", "[[target.md|Shown]]")
    config = replace(
        LinkRuntimeConfig(),
        non_embed_wiki_mode=wiki_mode,
        non_embed_markdown_mode="remove",
    )

    result = process_markdown_links(
        "![[child.md]]",
        source,
        link_config=config,
        target_format="docx",
    )

    if expected == "literal":
        assert result == escape_markdown_source_literal("[[target.md|Shown]]")
    elif expected == "text":
        assert result == "Shown"
    elif expected == "removed":
        assert result == ""
    else:
        assert result.startswith("[Shown](<")
        assert Path(target).resolve().as_uri() in result


def test_embed_boundary_policy_reaches_a_stable_point_after_cascade(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "a.md", "![[")
    _write(tmp_path / "mid.md", "[x](")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(
        "![[a.md]]mid.md]]https://evil.example)",
        source,
        link_config=config,
        target_format="docx",
    )

    assert "[x](https://evil.example)" not in result
    assert "{{IMAGE:" not in result


@pytest.mark.parametrize(
    ("mode", "expected"),
    [
        ("ignore", "before  after"),
        ("keep", "before ![[grand.md|Alias]] after"),
        ("placeholder", "before [Max depth reached: grand.md] after"),
    ],
)
def test_max_depth_policy_replaces_only_the_blocked_nested_embed(
    tmp_path: Path,
    mode: str,
    expected: str,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "child.md", "before ![[grand.md|Alias]] after")
    _write(tmp_path / "grand.md", "grand content")
    config = replace(
        LinkRuntimeConfig(),
        max_depth=1,
        max_depth_reached_mode=mode,
    )

    result = process_markdown_links("![[child.md]]", source, link_config=config)

    assert result == expected


def test_embedded_section_error_keep_preserves_original_path_and_alias(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "nested" / "child.md", "# Existing\nbody\n")
    text = "![[nested/child.md#Missing|Alias]]"
    config = replace(LinkRuntimeConfig(), file_not_found_mode="keep")

    result = process_markdown_links(text, source, link_config=config)

    assert result == text


def test_nested_yaml_close_requires_a_standalone_delimiter_line(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    _write(
        tmp_path / "child.md",
        '\ufeff---\ntitle: "a---b"\n---\nBODY\n',
    )

    result = process_markdown_links(
        "![[child.md]]",
        source,
        link_config=LinkRuntimeConfig(),
    )

    assert result == "BODY\n"


def test_invalid_mismatched_inline_code_run_does_not_hide_link_policy(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    text = "`[[missing.md|Shown]]`` tail"
    config = replace(LinkRuntimeConfig(), non_embed_wiki_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert "[[missing.md|Shown]]" not in result


@pytest.mark.parametrize(
    "text",
    [
        "paragraph\n    [[missing.md|Shown]]",
        "- item\n    [[missing.md|Shown]]",
    ],
)
def test_indented_paragraph_or_list_continuation_does_not_hide_link_policy(
    tmp_path: Path,
    text: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_wiki_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert "[[missing.md|Shown]]" not in result


@pytest.mark.parametrize(
    "prefix",
    [
        "# Heading\n",
        "---\n",
        "- - -\n",
        "Title\n===\n",
        "````markdown\ninside\n````\n",
    ],
)
def test_indented_code_after_block_boundary_hides_link_policy(
    tmp_path: Path,
    prefix: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"{prefix}    [[missing.md|Shown]]\n"
    config = replace(LinkRuntimeConfig(), non_embed_wiki_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert result == text


@pytest.mark.parametrize("prefix", ["\t", " \t", "  \t", "   \t"])
def test_top_level_tab_stop_indented_code_hides_link_policy(
    tmp_path: Path,
    prefix: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"{prefix}[[missing.md|Shown]]\n"
    config = replace(LinkRuntimeConfig(), non_embed_wiki_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert result == text


def test_all_link_rewrites_protect_fenced_and_inline_code(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    _write_png(tmp_path / "inside.png")
    _write_png(tmp_path / "outside.png")
    _write(tmp_path / "page.md", "page content\n")
    config = LinkRuntimeConfig.from_config(
        {
            "non_embed_links": {
                "wiki_mode": "extract_text",
                "markdown_mode": "extract_text",
                "auto_link_bare_url": True,
            },
            "embed_links": {
                "wiki_image_mode": "remove",
                "markdown_image_mode": "remove",
                "md_file_mode": "keep",
            },
        }
    )
    inline_code = (
        "`![[inside.png]] ![inside](inside.png) [[page.md|Inline]] "
        "[link](https://inline.example) https://bare-inline.example`"
    )
    fenced_code = (
        "```markdown\n"
        "![[inside.png]]\n"
        "![inside](inside.png)\n"
        "[[page.md|Fenced]]\n"
        "[link](https://fenced.example)\n"
        "https://bare-fenced.example\n"
        "```"
    )
    outside = (
        "![[outside.png]] ![outside](outside.png) [[page.md|Outside]] "
        "[site](https://linked.example) https://bare-outside.example"
    )

    result = process_markdown_links(
        f"{outside}\n{inline_code}\n{fenced_code}\n",
        source,
        link_config=config,
        target_format="docx",
    )

    assert inline_code in result
    assert fenced_code in result
    assert "![[outside.png]]" not in result
    assert "![outside](outside.png)" not in result
    assert "[[page.md|Outside]]" not in result
    assert "[site](https://linked.example)" not in result
    assert "<https://bare-outside.example>" in result


def test_target_aware_keep_and_bare_url_contract(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(
        LinkRuntimeConfig(),
        non_embed_wiki_mode="keep",
        non_embed_markdown_mode="keep",
        embed_markdown_image_mode="keep",
        auto_link_bare_url=True,
    )
    text = (
        "[site](https://linked.example) ![alt](image.png) "
        "[[https://wiki.example|wiki]] <https://angle.example> "
        "https://bare.example."
    )

    docx = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format="docx",
    )
    spreadsheet = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format="xlsx",
    )

    assert escape_markdown_source_literal("[site](https://linked.example)") in docx
    assert escape_markdown_source_literal("![alt](image.png)") in docx
    assert escape_markdown_source_literal("[[https://wiki.example|wiki]]") in docx
    assert "<https://angle.example>" in docx
    assert "<https://bare.example>." in docx
    assert spreadsheet == text


def test_recursive_embeds_do_not_rewrite_code_in_embedded_markdown(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    _write(
        tmp_path / "outer.md",
        "Outside ![[visible.md]]\n`![[inline.md]]`\n```md\n![[fenced.md]]\n```\n",
    )
    _write(tmp_path / "visible.md", "visible content\n")
    _write(tmp_path / "inline.md", "inline content must stay hidden\n")
    _write(tmp_path / "fenced.md", "fenced content must stay hidden\n")

    result = process_markdown_links(
        "![[outer.md]]",
        source,
        link_config=LinkRuntimeConfig(),
        target_format="docx",
    )

    assert "visible content" in result
    assert "`![[inline.md]]`" in result
    assert "```md\n![[fenced.md]]\n```" in result
    assert "inline content must stay hidden" not in result
    assert "fenced content must stay hidden" not in result


def test_table_safe_sized_image_placeholder_escapes_dimension_pipes(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    image = _write_png(tmp_path / "pixel.png")

    result = process_markdown_links(
        "| Image | Note |\n| --- | --- |\n| ![pixel](pixel.png =20x10) | ok |\n",
        source,
        link_config=LinkRuntimeConfig(),
        target_format="xlsx",
        table_safe=True,
    )

    assert f"{{{{IMAGE:{image}\\|20\\|10}}}}" in result
    assert f"{{{{IMAGE:{image}|20|10}}}}" not in result


def test_markdown_image_embed_accepts_commonmark_title(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    image = _write_png(tmp_path / "pixel.png")

    result = process_markdown_links(
        '![pixel](pixel.png "caption")',
        source,
        link_config=LinkRuntimeConfig(),
        target_format="docx",
    )

    assert result == f"{{{{IMAGE:{image}}}}}"


def test_link_rewrites_protect_strict_and_indented_code_blocks(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    fenced = (
        "````markdown\n"
        "[inside](https://inside.example)\n"
        "```not-a-closing-fence\n"
        "![inside](missing.png)\n"
        "https://bare-inside.example\n"
        "````"
    )
    indented = "    [indented](https://indented.example) ![image](missing.png)"
    config = replace(
        LinkRuntimeConfig(),
        non_embed_markdown_mode="remove",
        embed_markdown_image_mode="remove",
        auto_link_bare_url=True,
    )

    result = process_markdown_links(
        f"{fenced}\n{indented}\n",
        source,
        link_config=config,
        target_format="docx",
    )

    assert fenced in result
    assert indented in result


def test_invalid_backtick_info_string_does_not_open_a_code_fence(tmp_path: Path) -> None:
    source = _write(tmp_path / "invalid-fence.md")
    text = "```bad`info\n[remove me](https://remove.example)\n"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format="docx",
    )

    assert "```bad`info" in result
    assert "remove me" not in result
    assert "https://remove.example" not in result
