"""Declared WikiLink navigation targets avoid Vault/filesystem discovery."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import process_markdown_links
from docwen_core.links.declared_resources import (
    DeclaredResourceError,
    DeclaredResourceResolver,
    reject_declared_input_link_lookups,
)

pytestmark = pytest.mark.contract


def _bindings(source: str, token: str, href: str) -> dict:
    return {
        "authored_sha256": hashlib.sha256(source.encode()).hexdigest(),
        "images": [],
        "wiki_links": [{"authored_token": token, "href": href}],
    }


def test_declared_wiki_navigation_is_authenticated_and_needs_no_file_lookup(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    token = "[[Other#Section|Other note]]"
    source = f"See {token}.\n\n" + chr(96) + token + chr(96)
    href = "obsidian://open?vault=Knowledge&file=Notes%2FOther.md%23Section"
    resolver = DeclaredResourceResolver("Notes/Current.md", {}).with_bindings(
        source,
        _bindings(source, token, href),
    )

    def forbidden(*args: object, **kwargs: object) -> None:
        raise AssertionError("declared WikiLink attempted filesystem discovery")

    monkeypatch.setattr("docwen_core.links._non_embed.resolve_file_path", forbidden)
    reject_declared_input_link_lookups(
        source,
        wiki_mode="hyperlink",
        declared_wiki_link=resolver.resolve_wiki_link,
    )
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=LinkRuntimeConfig(non_embed_wiki_mode="hyperlink"),
        target_format="docx",
        declared_wiki_link=resolver.resolve_wiki_link,
    )

    assert "[Other note](<obsidian://open?vault=Knowledge&file=Notes%2FOther.md%23Section>)" in result
    assert chr(96) + token + chr(96) in result


@pytest.mark.parametrize(
    "source,token",
    [
        (chr(96) + "[[Other]]" + chr(96), "[[Other]]"),
        ("~~~md\n[[Other]]\n~~~", "[[Other]]"),
        ("Plain text", "[[Other]]"),
    ],
)
def test_declared_wiki_binding_requires_visible_authored_token(source: str, token: str) -> None:
    resolver = DeclaredResourceResolver("Notes/Current.md", {})
    with pytest.raises(DeclaredResourceError, match="visible authored wiki link"):
        resolver.with_bindings(
            source,
            _bindings(source, token, "obsidian://open?vault=Knowledge&file=Other.md"),
        )


def test_declared_wiki_binding_rejects_unsafe_target_scheme() -> None:
    source = "[[Other]]"
    resolver = DeclaredResourceResolver("Notes/Current.md", {})
    with pytest.raises(DeclaredResourceError, match="unsupported declared wiki link target"):
        resolver.with_bindings(source, _bindings(source, source, "javascript:alert(1)"))


def test_unbound_local_wiki_navigation_still_fails_closed() -> None:
    source = "[[Other]]"
    resolver = DeclaredResourceResolver("Notes/Current.md", {}).with_bindings(
        source,
        {
            "authored_sha256": hashlib.sha256(source.encode()).hexdigest(),
            "images": [],
            "wiki_links": [],
        },
    )
    with pytest.raises(DeclaredResourceError, match="local wiki link"):
        reject_declared_input_link_lookups(
            source,
            wiki_mode="hyperlink",
            declared_wiki_link=resolver.resolve_wiki_link,
        )


@pytest.mark.parametrize("mode", ["keep", "extract_text", "remove"])
def test_declared_navigation_does_not_override_text_only_wiki_modes(mode: str, tmp_path: Path) -> None:
    source = "[[Other|Label]]"
    resolver = DeclaredResourceResolver("Notes/Current.md", {}).with_bindings(
        source,
        _bindings(source, source, "obsidian://open?vault=Knowledge&file=Other.md"),
    )
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=LinkRuntimeConfig(non_embed_wiki_mode=mode),
        target_format="docx",
        declared_wiki_link=resolver.resolve_wiki_link,
    )

    if mode == "keep":
        assert result == r"\[\[Other\|Label\]\]"
    elif mode == "extract_text":
        assert result == "Label"
    else:
        assert result == ""
