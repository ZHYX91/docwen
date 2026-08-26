"""Focused tests split from test_link_runtime_processing_contract.py."""

from __future__ import annotations

from ._link_runtime_processing_contract_support import (
    LinkRuntimeConfig,
    Path,
    _write,
    _write_png,
    escape_markdown_source_literal,
    inspect,
    process_markdown_links,
    pytest,
    replace,
)

pytestmark = pytest.mark.contract


def test_from_config_includes_auto_link_bare_url() -> None:
    enabled = LinkRuntimeConfig.from_config({"non_embed_links": {"auto_link_bare_url": True}})
    malformed = LinkRuntimeConfig.from_config({"non_embed_links": {"auto_link_bare_url": "false"}})

    assert enabled.auto_link_bare_url is True
    assert malformed.auto_link_bare_url is False
    assert LinkRuntimeConfig().auto_link_bare_url is False


def test_link_config_drives_max_depth_and_max_depth_error_policy(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "outer.md", "outer\n![[inner.md]]\n")
    _write(tmp_path / "inner.md", "inner content\n")
    config = replace(
        LinkRuntimeConfig(),
        max_depth=1,
        max_depth_reached_mode="placeholder",
    )

    result = process_markdown_links(
        "![[outer.md]]",
        source,
        link_config=config,
        target_format="docx",
    )

    assert "Max depth reached" in result
    assert "inner content" not in result


def test_link_policy_has_one_public_input_object() -> None:
    parameters = inspect.signature(process_markdown_links).parameters
    dispersed_policy_parameters = {
        "max_depth",
        "image_mode",
        "markdown_image_mode",
        "md_mode",
        "wiki_mode",
        "markdown_mode",
        "auto_link_bare_url",
        "search_dirs",
        "on_not_found",
        "on_circular",
        "on_max_depth",
        "detect_circular",
    }

    assert parameters["link_config"].default is inspect.Parameter.empty
    assert dispersed_policy_parameters.isdisjoint(parameters)


def test_wiki_and_markdown_image_modes_are_independent(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    _write_png(tmp_path / "wiki.png")
    _write_png(tmp_path / "markdown.png")
    config = replace(
        LinkRuntimeConfig(),
        embed_wiki_image_mode="remove",
        embed_markdown_image_mode="extract_text",
    )
    text = "Wiki: ![[wiki.png|Wiki alt]]\nMarkdown: ![Markdown alt](markdown.png =20x10)"

    result = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format="docx",
    )

    assert "![[wiki.png|Wiki alt]]" not in result
    assert "Wiki alt" not in result
    assert "![Markdown alt](markdown.png =20x10)" not in result
    assert "Markdown: Markdown alt" in result


def test_markdown_image_embed_resolves_search_dirs_and_emits_placeholder(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    image = _write_png(tmp_path / "vault" / "markdown.png")
    config = replace(
        LinkRuntimeConfig(),
        embed_markdown_image_mode="embed",
        search_dirs=("vault",),
    )

    result = process_markdown_links(
        "![Markdown alt](markdown.png =20x10)",
        source,
        link_config=config,
        target_format="xlsx",
    )

    assert result == f"{{{{IMAGE:{image}|20|10}}}}"


def test_markdown_image_embed_uses_file_not_found_policy(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    original = "![Missing](missing.png)"

    kept = process_markdown_links(
        original,
        source,
        link_config=replace(LinkRuntimeConfig(), file_not_found_mode="keep"),
        target_format="xlsx",
    )
    placeholder = process_markdown_links(
        original,
        source,
        link_config=replace(LinkRuntimeConfig(), file_not_found_mode="placeholder"),
        target_format="xlsx",
    )

    assert kept == original
    assert placeholder == "[File not found: missing.png]"


def test_link_config_search_dirs_are_used_for_embed_resolution(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "vault" / "shared.md", "content from custom search directory\n")
    config = replace(LinkRuntimeConfig(), search_dirs=("vault",))

    result = process_markdown_links(
        "![[shared.md]]",
        source,
        link_config=config,
        target_format="docx",
    )

    assert "content from custom search directory" in result
    assert "File not found" not in result


def test_link_config_file_not_found_policy_is_consumed(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), file_not_found_mode="keep")

    result = process_markdown_links(
        "before ![[missing.md]] after",
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == ("before " + escape_markdown_source_literal("![[missing.md]]") + " after")


def test_link_config_circular_reference_policy_is_consumed(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "a.md", "A before\n![[b.md]]\nA after\n")
    _write(tmp_path / "b.md", "B before\n![[a.md]]\nB after\n")
    config = replace(
        LinkRuntimeConfig(),
        max_depth=10,
        circular_reference_mode="ignore",
    )

    result = process_markdown_links(
        "![[a.md]]",
        source,
        link_config=config,
        target_format="docx",
    )

    assert "A before" in result
    assert "B before" in result
    assert "Circular reference" not in result
    assert "![[a.md]]" not in result


def test_hyperlink_mode_uses_renderable_markdown_for_docx(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "page.md", "page content\n")
    config = replace(
        LinkRuntimeConfig(),
        non_embed_wiki_mode="hyperlink",
        non_embed_markdown_mode="hyperlink",
    )

    result = process_markdown_links(
        "See [[page.md|Page]] and [Example](https://example.com).",
        source,
        link_config=config,
        target_format="docx",
    )

    assert "[Page](<page.md>)" in result
    assert "[Example](<https://example.com>)" in result
    assert "[[page.md|Page]]" not in result


def test_wiki_hyperlink_is_not_reprocessed_by_markdown_keep(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "page.md", "page content\n")
    config = replace(
        LinkRuntimeConfig(),
        non_embed_wiki_mode="hyperlink",
        non_embed_markdown_mode="keep",
    )

    result = process_markdown_links(
        "[[page.md|Page]] and [Literal](https://example.com)",
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == ("[Page](<page.md>) and " + escape_markdown_source_literal("[Literal](https://example.com)"))


def test_fragment_only_wiki_hyperlink_downgrades_to_display_text(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_wiki_mode="hyperlink")

    result = process_markdown_links(
        "Jump to [[#section-name|Section]].",
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == "Jump to Section."


@pytest.mark.parametrize(
    ("mode", "expected"),
    [
        (
            "keep",
            escape_markdown_source_literal("[[missing.md#Part|Shown]]"),
        ),
        ("ignore", ""),
        ("placeholder", "[File not found: missing.md#Part]"),
    ],
)
def test_wiki_hyperlink_missing_target_consumes_error_policy(
    tmp_path: Path,
    mode: str,
    expected: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(
        LinkRuntimeConfig(),
        non_embed_wiki_mode="hyperlink",
        file_not_found_mode=mode,
    )

    result = process_markdown_links(
        "[[missing.md#Part|Shown]]",
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == expected


@pytest.mark.parametrize("target_format", ["xlsx", "csv"])
def test_docx_hyperlink_representation_does_not_leak_to_spreadsheets(
    tmp_path: Path,
    target_format: str,
) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "page.md", "page content\n")
    config = replace(
        LinkRuntimeConfig(),
        non_embed_wiki_mode="hyperlink",
        non_embed_markdown_mode="hyperlink",
    )
    text = "See [[page.md|Page]] and [Example](https://example.com)."

    result = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format=target_format,
    )

    assert result == text


@pytest.mark.parametrize(
    "text",
    [
        "-\n    [[missing.md|Shown]]",
        "=\n    [[missing.md|Shown]]",
    ],
)
def test_single_list_or_equals_line_does_not_hide_link_policy(
    tmp_path: Path,
    text: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_wiki_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert "[[missing.md|Shown]]" not in result
    assert "{{LINK:" not in result


def test_md_file_keep_preserves_original_nested_path_and_alias(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    _write(tmp_path / "nested" / "child.md", "child content\n")
    config = replace(LinkRuntimeConfig(), embed_md_file_mode="keep")
    text = "before ![[nested/child.md|Alias]] after"

    result = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == ("before " + escape_markdown_source_literal("![[nested/child.md|Alias]]") + " after")


@pytest.mark.parametrize("label", ["outer [inner]", r"outer \] bracket"])
def test_markdown_link_policy_handles_nested_or_escaped_label(
    tmp_path: Path,
    label: str,
) -> None:
    source = _write(tmp_path / "source.md")
    text = f"before [{label}](https://example.com) after"
    config = replace(
        LinkRuntimeConfig(),
        non_embed_markdown_mode="remove",
    )

    result = process_markdown_links(
        text,
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == "before  after"


@pytest.mark.parametrize(
    "target",
    [
        "https://example.com/O'Reilly",
        'https://example.com/foo"bar',
    ],
)
def test_markdown_link_policy_treats_inline_quotes_as_destination_text(
    tmp_path: Path,
    target: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(
        f"before [label]({target}) after",
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == "before  after"


@pytest.mark.parametrize(
    ("source_target", "expected_target"),
    [
        (r"https\://example.com/a\/b", "https://example.com/a/b"),
        (r"mailto\:user@example.com", "mailto:user@example.com"),
    ],
)
def test_markdown_destination_escapes_preserve_remote_uri_structure(
    tmp_path: Path,
    source_target: str,
    expected_target: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="hyperlink")

    result = process_markdown_links(
        f"[label]({source_target})",
        source,
        link_config=config,
        target_format="docx",
    )

    assert result == f"[label](<{expected_target}>)"


def test_escaped_remote_image_target_cannot_resolve_as_a_local_path(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), embed_markdown_image_mode="embed")

    result = process_markdown_links(
        r"![remote](https\://example.com/pixel.png)",
        source,
        link_config=config,
        target_format="docx",
    )

    assert "{{IMAGE" not in result
    assert "Remote embed fetching is unsupported" in result
    assert "File not found" not in result


@pytest.mark.parametrize(
    "text",
    [
        "before [`code`](https://example.com) after",
        "before [label](https://example.com/`code`) after",
    ],
)
def test_markdown_link_policy_processes_code_spans_inside_construct(
    tmp_path: Path,
    text: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert result == "before  after"


def test_markdown_image_policy_processes_code_span_inside_alt(
    tmp_path: Path,
) -> None:
    source = _write(tmp_path / "source.md")
    text = "before ![`alt`](image.png) after"
    config = replace(LinkRuntimeConfig(), embed_markdown_image_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert result == "before  after"


def test_wiki_policy_processes_code_span_inside_display(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    text = "before [[target.md|`code`]] after"
    config = replace(LinkRuntimeConfig(), non_embed_wiki_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert result == "before  after"


def test_escaped_backticks_do_not_hide_markdown_link_policy(tmp_path: Path) -> None:
    source = _write(tmp_path / "source.md")
    text = r"\`[label](https://example.com)\`"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(text, source, link_config=config)

    assert "[label](https://example.com)" not in result


@pytest.mark.parametrize(
    "syntax",
    [
        "[x](foo bar)",
        "[x](<foo>bar)",
        r"[x](<a\#b>)",
        '[x](foo "title"junk)',
        r"""[x](foo "a\qb")""",
        "[x\n\n y](url)",
        "[x\r\r y](url)",
        "[x](foo\n\nbar)",
    ],
)
def test_invalid_markdown_link_syntax_is_not_treated_as_a_policy_target(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(syntax, source, link_config=config)

    assert result == syntax


@pytest.mark.parametrize(
    "syntax",
    [
        "[x](\u00a0foo)",
        "[x](foo\u00a0bar)",
        "[x](foo\u2003bar)",
    ],
)
def test_non_ascii_whitespace_remains_part_of_a_link_destination(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(syntax, source, link_config=config)

    assert result == ""


@pytest.mark.parametrize(
    "syntax",
    [
        "[x\r\ny](url)",
        '[x](url\r\n"title")',
    ],
)
def test_single_crlf_inside_a_valid_link_is_not_a_blank_line(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = _write(tmp_path / "source.md")
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    assert process_markdown_links(syntax, source, link_config=config) == ""
