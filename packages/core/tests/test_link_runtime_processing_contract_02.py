"""Focused tests split from test_link_runtime_processing_contract.py."""

from __future__ import annotations

from ._link_runtime_processing_contract_support import (
    LinkRuntimeConfig,
    Path,
    _write,
    process_markdown_links,
    pytest,
    replace,
)

pytestmark = pytest.mark.contract


def test_reference_tail_is_skipped_only_when_its_definition_exists(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    defined = process_markdown_links(
        "[bar]: /ref\n\n[foo][bar](baz)",
        source,
        link_config=config,
    )
    undefined = process_markdown_links(
        "[foo][bar](baz)",
        source,
        link_config=config,
    )
    escaped = process_markdown_links(
        "[foo][bar\\]](baz)\n\n[bar\\]]: /ref",
        source,
        link_config=config,
    )
    invalid = process_markdown_links(
        "[foo][bar](baz)\n\n[bar]:",
        source,
        link_config=config,
    )

    assert defined == "[bar]: /ref\n\n[foo][bar](baz)"
    assert undefined == "[foo]"
    assert escaped == "[foo][bar\\]](baz)\n\n[bar\\]]: /ref"
    assert invalid == "[foo]\n\n[bar]:"


@pytest.mark.parametrize("definition", ["> [bar]: /ref", "- [bar]: /ref"])
def test_container_reference_definition_owns_the_following_reference_tail(
    tmp_path: Path,
    definition: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"{definition}\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize(
    "definition",
    ["paragraph\n[bar]: /ref", "- item\n  [bar]: /ref"],
)
def test_paragraph_continuation_is_not_a_reference_definition(
    tmp_path: Path,
    definition: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"{definition}\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert result == f"{definition}\n\n[foo]"


def test_blockquote_paragraph_continuation_is_not_a_reference_definition(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    definition = "> paragraph\n> [bar]: /ref"
    text = f"{definition}\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert result == f"{definition}\n\n[foo]"


def test_empty_footnote_definition_owns_the_footnote_reference_tail(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    text = "[foo][^bar](baz)\n\n[^bar]:"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_collapsed_reference_previous_label_ignores_escaped_open_bracket(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    text = r"[foo \[x][](baz)" + "\n\n" + r"[foo \[x]: /ref"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize(("label_length", "defined"), [(500, True), (501, False)])
def test_reference_definition_label_limit_matches_extended_parser(
    tmp_path: Path,
    label_length: int,
    defined: bool,
) -> None:
    source = _write(tmp_path / "source.md")
    label = "a" * label_length
    definition = f"[{label}]: /ref"
    text = f"{definition}\n\n[foo][{label}](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    expected = text if defined else f"{definition}\n\n[foo]"
    assert result == expected


@pytest.mark.parametrize(
    "first_definition",
    ['[a]: /a\n  "title"', "[a]:\n /a"],
)
def test_reference_definition_continuation_keeps_following_definition_global(
    tmp_path: Path,
    first_definition: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"{first_definition}\n[b]: /b\n\n[foo][b](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize("block", ["# heading", "---", "* * *"])
def test_completed_heading_or_thematic_block_allows_reference_definition(
    tmp_path: Path,
    block: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"{block}\n[bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_footnote_definition_rejects_unescaped_whitespace_key(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    text = "[foo][^a b](baz)\n\n[^a b]:"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == "[foo]\n\n[^a b]:"


def test_empty_footnote_body_owns_immediately_following_definition_syntax(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    text = "[^a]:\n[bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == "[^a]:\n[bar]: /u\n\n[foo]"


def test_nested_blockquote_transition_can_start_a_global_reference_definition(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    text = "> paragraph\n> > [bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_escaped_reference_label_limit_counts_parser_units(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    label = r"\]" * 251
    text = f"[{label}]: /u\n\n[foo][{label}](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_escaped_footnote_label_limit_counts_parser_units(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    key = r"\]" * 251
    text = f"[foo][^{key}](baz)\n\n[^{key}]:"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize("non_blank", ["\u00a0", "\u2003", "\x85", "\u2028"])
def test_unicode_whitespace_only_line_does_not_interrupt_a_paragraph_for_definition(
    tmp_path: Path,
    non_blank: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"paragraph\n{non_blank}\n[bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == (f"paragraph\n{non_blank}\n[bar]: /u\n\n[foo]")


@pytest.mark.parametrize("quote", ['"', "'"])
def test_reference_definition_title_rejects_nul(
    tmp_path: Path,
    quote: str,
) -> None:
    source = _write(tmp_path / "source.md")
    definition = f"[bar]: /u {quote}\x00{quote}"
    text = f"{definition}\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == (f"{definition}\n\n[foo]")


@pytest.mark.parametrize("footnote", ["[^a]: text", "[^a]: "])
def test_nonempty_or_space_only_footnote_line_does_not_consume_next_definition(
    tmp_path: Path,
    footnote: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"{footnote}\n[bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize("indent", [" ", "    "])
def test_empty_footnote_consumes_one_body_line_then_reopens_definition_boundary(
    tmp_path: Path,
    indent: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"[^a]:\n{indent}body\n[bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize("footnote", ["[^a]:\nbody", "[^a]: text"])
def test_footnote_continuation_keeps_following_definition_global(
    tmp_path: Path,
    footnote: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"{footnote}\n  continuation\n   repeated\n[bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize("continuation", ["  continuation", " \tcontinuation", "   \vcontinuation"])
def test_footnote_continuation_survives_blank_and_parser_whitespace(
    tmp_path: Path,
    continuation: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"[^a]: text\n\n{continuation}\n[bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_footnote_continuation_indent_is_relative_to_definition_lead(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    text = "   [^a]: text\n\n       continuation\n[bar]: /u\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_footnote_definition_can_interrupt_a_paragraph(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    text = "paragraph\n[^a]: text\n\n[foo][^a](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize("escaped", [r"a\q", "a\\ "])
def test_footnote_label_allows_any_escaped_character(
    tmp_path: Path,
    escaped: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"[foo][^{escaped}](baz)\n\n[^{escaped}]:"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_generic_reference_escaped_non_punctuation_counts_as_one_unit(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    label = r"\q" * 251
    text = f"[{label}]: /u\n\n[foo][{label}](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_decreasing_blockquote_depth_does_not_interrupt_outer_paragraph(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    definition = "> > paragraph\n> [bar]: /u"
    text = f"{definition}\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == (f"{definition}\n\n[foo]")


@pytest.mark.parametrize("non_blank", ["\u00a0", "\u2003"])
def test_unicode_quote_line_does_not_create_a_definition_boundary(
    tmp_path: Path,
    non_blank: str,
) -> None:
    source = _write(tmp_path / "source.md")
    definition = f"> paragraph\n> {non_blank}\n> [bar]: /u"
    text = f"{definition}\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == (f"{definition}\n\n[foo]")


@pytest.mark.parametrize("definition", ["[^a b]: /u", "[^a]:/u"])
def test_invalid_footnote_syntax_can_fall_back_to_normal_reference_definition(
    tmp_path: Path,
    definition: str,
) -> None:
    source = _write(tmp_path / "source.md")
    label = definition.split("]:", 1)[0][1:]
    text = f"{definition}\n\n[foo][{label}](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


def test_bare_reference_destination_accepts_nul_like_extended_parser(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    definition = "[bar]: u\x00"
    text = f"{definition}\n\n[foo][bar](baz)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(text, source, link_config=config) == text


@pytest.mark.parametrize("target", ['u "\x00"', "u '\x00'", '<a> "\x00"'])
def test_inline_link_title_rejects_nul_like_extended_parser(
    tmp_path: Path,
    target: str,
) -> None:
    source = _write(tmp_path / "source.md")
    syntax = f"[x]({target})"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(syntax, source, link_config=config) == syntax


@pytest.mark.parametrize(
    "syntax",
    [
        "[x\n# heading](url)",
        "[x\n> quote](url)",
        "[x\n- item](url)",
        "[x\n```md](url)",
        "[x\n<div>y](url)",
        "[x\n<!-- y -->](url)",
    ],
)
def test_link_scanner_does_not_cross_a_markdown_block_interrupt(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(syntax, source, link_config=config) == syntax


@pytest.mark.parametrize("syntax", ["[x\n---](url)", "[x\n=](url)"])
def test_marker_text_inside_a_valid_multiline_label_is_processed(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(syntax, source, link_config=config) == ""


def test_single_character_angle_destination_is_a_valid_policy_target(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links("[x](<a>)", source, link_config=config) == ""


@pytest.mark.parametrize("destination", ["<a<b>", "<a\x00b>"])
def test_invalid_angle_destination_characters_are_not_skipped_by_scanner(
    tmp_path: Path,
    destination: str,
) -> None:
    source = _write(tmp_path / "source.md")
    syntax = f"[x]({destination})"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(syntax, source, link_config=config) == syntax


@pytest.mark.parametrize("backslash_count", [2, 4])
def test_even_backslash_run_before_nested_label_close_matches_mistune(
    tmp_path: Path,
    backslash_count: int,
) -> None:
    source = _write(tmp_path / "source.md")
    syntax = "[[]" + ("\\" * backslash_count) + "](url)"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(syntax, source, link_config=config) == syntax


def test_escaped_literal_backslash_is_percent_protected_in_docx_target(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="hyperlink")

    result = process_markdown_links(
        r"[x](foo\\bar)",
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == "[x](<foo%5Cbar>)"


@pytest.mark.parametrize(
    "url",
    [
        "https://example.com/O'Reilly",
        'https://example.com/foo"bar',
        "https://example.com/[x]",
        "https://example.com/a_(b)",
        "https://example.com/a*b*",
    ],
)
def test_bare_url_policy_uses_a_safe_autolink(tmp_path: Path, url: str) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), auto_link_bare_url=True)

    result = process_markdown_links(
        url,
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == f"<{url}>"


@pytest.mark.parametrize(
    "syntax",
    [
        '<span title="[x](https://example.com)">ok</span>',
        '<span data="[[x|y]]">ok</span>',
        "<https://host.example/[x](y)>",
        "$[x](https://example.com)$",
        "$$[x](https://example.com)$$",
        "$$\n[x](https://example.com)\n$$",
        "<foo>\n[x](url) [[x|y]]\n</foo>",
        "> ~~~\n> [x](url) [[x|y]]\n> ~~~",
        "- ~~~\n  [x](url) [[x|y]]\n  ~~~",
        ">     [x](url) [[x|y]]",
        "> <div>\n> [x](url) [[x|y]]\n> </div>",
        "[[x `]` y]]",
        "![[x `]` y]]",
    ],
)
def test_link_policy_preserves_renderer_atoms_and_container_blocks(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(
        LinkRuntimeConfig(),
        non_embed_markdown_mode="remove",
        non_embed_wiki_mode="remove",
        embed_wiki_image_mode="remove",
    )

    result = process_markdown_links(
        syntax,
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == syntax


@pytest.mark.parametrize(
    "url",
    [
        "https://host.example/[x](y)",
        "https://host.example/![x](y.png)",
        "https://host.example/[[x|y]]",
        "https://host.example/`code`more",
        "https://host.example/$math$more",
    ],
)
def test_bare_url_precedence_protects_link_like_url_content(
    tmp_path: Path,
    url: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(
        LinkRuntimeConfig(),
        non_embed_markdown_mode="remove",
        non_embed_wiki_mode="remove",
        embed_markdown_image_mode="remove",
        auto_link_bare_url=True,
    )

    result = process_markdown_links(
        url,
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == f"<{url}>"


@pytest.mark.parametrize(
    "syntax",
    [
        "$$\n[x](url)\n $$",
        "$$\n[x](url)",
        "<!doctype\n[x](url)\n>",
        "<![cdata[\n[x](url)\n]]>",
    ],
)
def test_atom_protection_does_not_exceed_mistune_grammar(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(syntax, source, link_config=config)

    assert "[x](url)" not in result
