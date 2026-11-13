"""Independent request-scoped ordinary link modes and selected fallbacks."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import process_markdown_links
from docwen_core.links._markdown_inline import escape_markdown_source_literal

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("target_format", ["docx", "xlsx", "csv"])
@pytest.mark.parametrize(
    "syntax,address",
    [
        ("[Label](<nested/page.md#Heading>)", "nested/page.md#Heading"),
        ("[[nested/page.md#Heading|Label]]", "nested/page.md#Heading"),
    ],
)
def test_extract_address_never_expands_to_private_absolute_path(
    tmp_path: Path, target_format: str, syntax: str, address: str
) -> None:
    config = replace(
        LinkRuntimeConfig(), non_embed_wiki_mode="extract_address", non_embed_markdown_mode="extract_address"
    )
    result = process_markdown_links(
        syntax, str(tmp_path / "source.md"), link_config=config, target_format=target_format
    )
    expected = escape_markdown_source_literal(address) if target_format == "docx" else address
    assert result == expected
    assert str(tmp_path) not in result


@pytest.mark.parametrize("target_format", ["docx", "xlsx", "csv"])
@pytest.mark.parametrize("fallback", ["keep", "extract_text", "extract_address"])
@pytest.mark.parametrize(
    "syntax,address",
    [
        ("[Label](javascript:alert)", "javascript:alert"),
        ("[[#missing|Label]]", "#missing"),
        ("[Label](missing.md)", "missing.md"),
    ],
)
def test_unsupported_or_missing_target_uses_independent_fallback(
    tmp_path: Path, target_format: str, fallback: str, syntax: str, address: str
) -> None:
    config = replace(LinkRuntimeConfig(), wiki_hyperlink_fallback=fallback, markdown_hyperlink_fallback=fallback)
    result = process_markdown_links(
        syntax, str(tmp_path / "source.md"), link_config=config, target_format=target_format
    )
    expected = {"keep": syntax, "extract_text": "Label", "extract_address": address}[fallback]
    if target_format == "docx" and fallback != "extract_text":
        expected = escape_markdown_source_literal(expected)
    assert result == expected


def test_fallbacks_do_not_cross_syntax_boundary(tmp_path: Path) -> None:
    config = replace(
        LinkRuntimeConfig(), wiki_hyperlink_fallback="extract_text", markdown_hyperlink_fallback="extract_address"
    )
    result = process_markdown_links(
        "[[#missing|Wiki]] [MD](#missing)", str(tmp_path / "source.md"), link_config=config, target_format="csv"
    )
    assert result == "Wiki #missing"


def test_extract_address_excludes_title_and_preserves_escaping(tmp_path: Path) -> None:
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="extract_address")
    result = process_markdown_links(
        '[Label](<a%23b.md#part> "Not an address")',
        str(tmp_path / "source.md"),
        link_config=config,
        target_format="csv",
    )
    assert result == "a%23b.md#part"


@pytest.mark.parametrize(
    "target",
    [
        "",
        "http:relative",
        "https://",
        "obsidian:open",
        "mailto:",
        "file:",
        "javascript:alert(1)",
        "data:text/plain,hello",
    ],
)
def test_non_navigable_uri_defaults_to_display_text(tmp_path: Path, target: str) -> None:
    original = f"[Label](<{target}>)"
    result = process_markdown_links(
        original, str(tmp_path / "source.md"), link_config=LinkRuntimeConfig(), target_format="docx"
    )
    assert result == "Label"


def test_declared_link_boundary_does_not_probe_markdown_target(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    def forbidden(*args, **kwargs):
        raise AssertionError("target filesystem lookup is forbidden")

    monkeypatch.setattr(Path, "is_file", forbidden)
    result = process_markdown_links(
        "[Label](private.md)",
        str(tmp_path / "source.md"),
        link_config=LinkRuntimeConfig(),
        target_format="docx",
        local_lookup_allowed=False,
    )
    assert result == "Label"


@pytest.mark.parametrize("mode", ["keep", "extract_text", "extract_address", "remove"])
def test_declared_wiki_non_resolving_modes_do_not_require_resources(mode: str) -> None:
    from docwen_core.links import reject_declared_input_link_lookups

    reject_declared_input_link_lookups("[[private.md|Label]]", wiki_mode=mode)
