"""Declared image mappings preserve authored syntax and never discover local files."""

from __future__ import annotations

import base64
import hashlib
from dataclasses import replace
from pathlib import Path

import pytest

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import markdown_source_projection, process_markdown_links
from docwen_core.links._markdown_references import markdown_reference_definitions
from docwen_core.links.declared_resources import (
    DeclaredResourceError,
    DeclaredResourceResolver,
    bind_declared_markdown_images,
    reject_declared_input_link_lookups,
)

pytestmark = pytest.mark.contract


def _bindings(source: str, token: str, logical: str = "assets/chart.png") -> dict:
    return {
        "authored_sha256": hashlib.sha256(source.encode()).hexdigest(),
        "images": [{"authored_token": token, "logical_path": logical}],
    }


@pytest.mark.parametrize(
    "source,token",
    [
        ("`![[chart.png]]`", "![[chart.png]]"),
        ("```md\n![[chart.png]]\n```", "![[chart.png]]"),
        ("![chart](chart.png)", "![[chart.png]]"),
    ],
)
def test_rejects_bindings_for_inert_or_absent_images(source: str, token: str) -> None:
    resolver = DeclaredResourceResolver("notes/note.md", {"assets/chart.png": "/declared.png"})
    with pytest.raises(DeclaredResourceError, match="visible authored image"):
        resolver.with_bindings(source, _bindings(source, token))


def test_rejects_changed_source_undeclared_and_conflicting_bindings() -> None:
    source = "![[chart.png]]"
    resolver = DeclaredResourceResolver(
        "notes/note.md", {"assets/chart.png": "/declared.png", "other/chart.png": "/other.png"}
    )
    with pytest.raises(DeclaredResourceError, match="hash mismatch"):
        resolver.with_bindings(source + "changed", _bindings(source, source))
    with pytest.raises(DeclaredResourceError, match="undeclared"):
        resolver.with_bindings(source, _bindings(source, source, "private/chart.png"))
    bindings = _bindings(source, source)
    bindings["images"].append({"authored_token": source, "logical_path": "other/chart.png"})
    with pytest.raises(DeclaredResourceError, match="conflicting"):
        resolver.with_bindings(source, bindings)


@pytest.mark.parametrize(
    "wiki_mode,markdown_mode,expected",
    [
        ("embed", "remove", "IMAGE@"),
        ("extract_text", "remove", r"chart\.png"),
        ("remove", "extract_text", "Markdown"),
        ("keep", "remove", r"chart\.png"),
    ],
)
def test_declared_images_keep_distinct_policies_without_path_search(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    wiki_mode: str,
    markdown_mode: str,
    expected: str,
) -> None:
    image = tmp_path / "opaque.png"
    image.write_bytes(
        base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNUSFjAwMDAxAAGAA0qASTlOPBgAAAAAElFTkSuQmCC"
        )
    )
    wiki, markdown = "![[chart.png|120x80]]", "![Markdown](../assets/chart.png)"
    source = f"{wiki}\n\n{markdown}\n\n`![[missing.png]]`"
    bindings = _bindings(source, wiki)
    bindings["images"].append({"authored_token": markdown, "logical_path": "assets/chart.png"})
    resolver = DeclaredResourceResolver("notes/note.md", {"assets/chart.png": str(image)}).with_bindings(
        source, bindings
    )

    def forbidden(*args: object, **kwargs: object) -> None:
        raise AssertionError("declared input attempted filesystem discovery")

    monkeypatch.setattr("docwen_core.links._embed_dispatch.resolve_file_path", forbidden)
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=LinkRuntimeConfig(
            embed_wiki_image_mode=wiki_mode,
            embed_markdown_image_mode=markdown_mode,
        ),
        target_format="docx",
        image_scope="scope",
        declared_image=resolver.resolve_image,
    )
    assert expected in result
    assert "`![[missing.png]]`" in result
    if wiki_mode == "embed":
        assert "|120|80" in result
    else:
        assert "opaque.png" not in result


@pytest.mark.parametrize("mode", ["keep", "extract_text", "remove"])
def test_text_only_wiki_policies_need_no_file_lookup(mode: str) -> None:
    reject_declared_input_link_lookups("[[outside.md|Label]]", wiki_mode=mode)


def test_only_active_local_wiki_navigation_is_rejected() -> None:
    reject_declared_input_link_lookups("[Label](outside.md) `[[outside]]` ![[chart.png]] [[#heading]]")
    with pytest.raises(DeclaredResourceError, match="local wiki link"):
        reject_declared_input_link_lookups("[[outside.md]]")


@pytest.mark.parametrize(
    "token,definition,alt",
    [
        ("![ref][image]", "[image]: active.png\n", "ref"),
        ("![image][]", "[image]: active.png\n", "image"),
        ("![image]", "[image]: active.png\n", "image"),
        ("![ref][IMAGE   NAME]", '[image\nname]:\n <active.png>\n "title"\n', "ref"),
        ("![ref][image]", '[image]: active.png "multi\nline title"\n', "ref"),
        (r"![ref][im\]age]", "[im\\]age]: active.png\n", "ref"),
        ("![image]", "[^note]: text\n[image]: active.png\n", "image"),
        ("![image]", "[^note]:\n    body\n[image]: active.png\n", "image"),
    ],
)
@pytest.mark.parametrize("eol", ["\n", "\r\n"])
def test_reference_images_are_authenticated_and_follow_the_markdown_image_policy(
    token: str, definition: str, alt: str, eol: str, tmp_path: Path
) -> None:
    source = (token + "\n\n" + definition).replace("\n", eol)
    resolver = DeclaredResourceResolver("notes/note.md", {"assets/chart.png": "/declared.png"})
    bound = resolver.with_bindings(source, _bindings(source, token))
    assert bound.resolve_image(token, "active.png") == "/declared.png"
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), embed_markdown_image_mode="extract_text"),
    )
    assert result == source.replace(token, alt, 1)


@pytest.mark.parametrize(
    "source,token",
    [
        ("![ref][missing]\n\n[image]: active.png\n", "![ref][missing]"),
        ("![image]\n\n```md\n[image]: active.png\n```\n", "![image]"),
        ("![image]\n\n<span>\n[image]: active.png\n</span>\n", "![image]"),
        ("![image]\n\ntext\n[image]: active.png\n", "![image]"),
        ("`![image]`\n\n[image]: active.png\n", "![image]"),
        ("<span>\n![image]\n</span>\n\n[image]: active.png\n", "![image]"),
        ("%% ![image] %%\n\n[image]: active.png\n", "![image]"),
        ("![image]\n\n%%\n[image]: active.png\n%%\n", "![image]"),
        ('[other]: active.png "![image]"\n[image]: active.png\n', "![image]"),
        ("![image]\n\n[^note]:\n[image]: active.png\n", "![image]"),
        *[
            (f"![image]\n\nparagraph\n{non_blank}\n[image]: active.png\n", "![image]")
            for non_blank in ["\u00a0", "\u2003", "\x85", "\u2028"]
        ],
    ],
)
def test_reference_image_bindings_cannot_authenticate_inert_occurrences(source: str, token: str) -> None:
    resolver = DeclaredResourceResolver("notes/note.md", {"assets/chart.png": "/declared.png"})
    with pytest.raises(DeclaredResourceError, match="visible authored image"):
        resolver.with_bindings(source, _bindings(source, token))


def test_reference_definition_first_wins_without_promoting_invalid_titles() -> None:
    source = '[image]: first.png\n "not a complete title\n[IMAGE]: second.png\n'
    definitions = markdown_reference_definitions(source)
    assert definitions["IMAGE"].target == "first.png"


def test_reference_bindings_preserve_hash_inventory_and_conflict_checks() -> None:
    source = "![image]\n\n[image]: active.png\n"
    resolver = DeclaredResourceResolver("notes/note.md", {"assets/chart.png": "/declared.png"})
    with pytest.raises(DeclaredResourceError, match="hash mismatch"):
        resolver.with_bindings(source + "changed", _bindings(source, "![image]"))
    with pytest.raises(DeclaredResourceError, match="undeclared"):
        resolver.with_bindings(source, _bindings(source, "![image]", "outside.png"))


@pytest.mark.parametrize(
    "target,filename",
    [
        ("active&#46;png", "active.png"),
        ("active&period;png", "active.png"),
        ("active%20file.png", "active file.png"),
        ("active&amp;period;png", "active&period;png"),
        (r"active\&period;png", "active&period;png"),
        (r"active\%20file.png", "active%20file.png"),
        ("active%2520file.png", "active%20file.png"),
    ],
)
def test_reference_image_destination_decodes_escapes_entities_and_uri_once(
    target: str, filename: str, tmp_path: Path
) -> None:
    source = f"![][image]\n\n[image]: {target}\n"
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), embed_markdown_image_mode="extract_text"),
    )
    assert result == source.replace("![][image]", filename, 1)


def test_declared_binding_helper_uses_the_visible_definition_after_a_comment() -> None:
    hidden = "%%\n\n[image]: hidden.png\n\n![image]\n%%\n\n"
    source = hidden + "![image]\n\n[image]: visible.png\n"
    resolver = DeclaredResourceResolver("note.md", {"visible.png": "/visible.png", "hidden.png": "/hidden.png"})
    assert bind_declared_markdown_images(source, resolver) == (
        hidden + "![image](</visible.png>)\n\n[image]: visible.png\n"
    )


@pytest.mark.parametrize("eol", ["\n", "\r\n"])
def test_container_image_binding_preserves_the_raw_token_and_label_coordinates(eol: str) -> None:
    source = "> ![im\n> age]\n>\n> [im\n> age]: active.png\n".replace("\n", eol)
    token = "![im\n> age]".replace("\n", eol)
    resolver = DeclaredResourceResolver("note.md", {"assets/chart.png": "/declared.png"})
    assert resolver.with_bindings(source, _bindings(source, token)).resolve_image(token, "unused") == "/declared.png"
    projection = markdown_source_projection(source)
    close = source.index("age]") + 3
    assert len(projection) == len(source)
    assert projection[close - 3 : close] == "age"
    assert projection[close] == " "


@pytest.mark.parametrize(
    "token",
    ["![im\n\nage]", "![im\n- age]", "![im\n> > age]"],
)
def test_reference_image_binding_cannot_cross_an_inline_container_boundary(token: str) -> None:
    source = "> " + token + "\n\n[im age]: active.png\n"
    resolver = DeclaredResourceResolver("note.md", {"assets/chart.png": "/declared.png"})
    with pytest.raises(DeclaredResourceError, match="visible authored image"):
        resolver.with_bindings(source, _bindings(source, token))
