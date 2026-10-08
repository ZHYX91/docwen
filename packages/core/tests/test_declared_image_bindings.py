"""Declared image mappings preserve authored syntax and never discover local files."""

from __future__ import annotations

import base64
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
