"""VIS-168 frozen Link parser/path/POSIX completion contract."""

from __future__ import annotations

import os
from dataclasses import replace
from pathlib import Path

import pytest
from PIL import Image

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import process_markdown_links

pytestmark = pytest.mark.contract


def _source(tmp_path: Path) -> str:
    path = tmp_path / "source.md"
    path.write_text("source\n", encoding="utf-8")
    return str(path)


def _write_png(path: Path) -> None:
    with Image.new("RGB", (1, 1), "white") as image:
        image.save(path, format="PNG")


@pytest.mark.parametrize(
    ("text", "protected"),
    [
        (
            "> ```markdown\n> [inside](https://inside.example)\n> ```\n[outside](https://outside.example)\n",
            "> ```markdown\n> [inside](https://inside.example)\n> ```",
        ),
        (
            "- ~~~markdown\n  [inside](https://inside.example)\n  ~~~\n- [outside](https://outside.example)\n",
            "- ~~~markdown\n  [inside](https://inside.example)\n  ~~~",
        ),
        (
            "> - ```markdown\n>   [inside](https://inside.example)\n>   ```\n"
            "> [outside](https://outside.example)\n"
            "[top](https://top.example)\n",
            "> - ```markdown\n>   [inside](https://inside.example)\n>   ```",
        ),
        (
            "> ```markdown\n> [inside](https://inside.example)\n[outside](https://outside.example)\n",
            "> ```markdown\n> [inside](https://inside.example)",
        ),
    ],
)
def test_container_owned_fences_protect_only_their_source_extent(
    tmp_path: Path,
    text: str,
    protected: str,
) -> None:
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(
        text,
        _source(tmp_path),
        link_config=config,
        target_format="docx",
    )

    assert protected in result
    assert "inside.example" in result
    assert "outside.example" not in result
    assert "top.example" not in result


def test_legacy_custom_error_text_keys_are_ignored_compatibly() -> None:
    baseline = LinkRuntimeConfig.from_config({})
    with_legacy_keys = LinkRuntimeConfig.from_config(
        {
            "error_handling": {
                "file_not_found_text": "FILE {filename}",
                "circular_text": "CIRCULAR {filename}",
                "max_depth_text": "DEPTH {filename}",
            }
        }
    )

    assert with_legacy_keys == baseline
    assert not hasattr(with_legacy_keys, "file_not_found_text")
    assert not hasattr(with_legacy_keys, "circular_text")
    assert not hasattr(with_legacy_keys, "max_depth_text")


def test_embedded_markdown_frontmatter_is_never_link_processed(
    tmp_path: Path,
) -> None:
    source = _source(tmp_path)
    child = tmp_path / "child.md"
    child.write_text(
        "\ufeff---\nsecret: '![[must-not-expand.md]]'\n---\nBody ![[visible.md]]\n",
        encoding="utf-8",
    )
    (tmp_path / "must-not-expand.md").write_text("SECRET EXPANSION\n", encoding="utf-8")
    (tmp_path / "visible.md").write_text("VISIBLE EXPANSION\n", encoding="utf-8")

    result = process_markdown_links(
        "![[child.md]]",
        source,
        link_config=LinkRuntimeConfig(),
        target_format="docx",
    )

    assert "secret:" not in result
    assert "SECRET EXPANSION" not in result
    assert "VISIBLE EXPANSION" in result


def test_unclosed_frontmatter_delimiter_remains_processable_markdown(
    tmp_path: Path,
) -> None:
    text = "---\n[remove](https://visible.example)\n"
    config = replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove")

    result = process_markdown_links(
        text,
        _source(tmp_path),
        link_config=config,
        target_format="docx",
    )

    assert result == "---\n\n"


@pytest.mark.parametrize(
    "syntax",
    [
        "![remote](https://example.com/image.png)",
        "![remote](http://example.com/image.png)",
        "![remote](//example.com/image.png)",
        "![[https://example.com/remote.md]]",
    ],
)
def test_remote_embed_is_explicitly_unsupported_without_local_resolution(
    tmp_path: Path,
    syntax: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fail_resolve(*_args: object, **_kwargs: object) -> None:
        pytest.fail("remote embed reached the local filesystem resolver")

    monkeypatch.setattr(
        "docwen_core.links._embed_dispatch.resolve_file_path",
        fail_resolve,
    )

    result = process_markdown_links(
        syntax,
        _source(tmp_path),
        link_config=LinkRuntimeConfig(),
        target_format="docx",
    )

    assert "Remote embed fetching is unsupported" in result
    assert "example.com" in result
    assert "File not found" not in result


@pytest.mark.parametrize(
    ("mode", "expected"),
    [
        ("keep", "![remote](https://example.com/image.png)"),
        ("extract_text", "remote"),
        ("remove", ""),
    ],
)
def test_remote_image_non_embed_modes_never_fetch(
    tmp_path: Path,
    mode: str,
    expected: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fail_resolve(*_args: object, **_kwargs: object) -> None:
        pytest.fail("non-embed remote mode reached the local filesystem resolver")

    monkeypatch.setattr(
        "docwen_core.links._embed_dispatch.resolve_file_path",
        fail_resolve,
    )
    config = replace(LinkRuntimeConfig(), embed_markdown_image_mode=mode)

    result = process_markdown_links(
        "![remote](https://example.com/image.png)",
        _source(tmp_path),
        link_config=config,
        target_format="xlsx",
    )

    assert result == expected


@pytest.mark.skipif(os.name == "nt", reason="POSIX permits a literal pipe in a filename")
def test_posix_percent_encoded_pipe_resolves_literal_pipe_file(
    tmp_path: Path,
) -> None:
    source = _source(tmp_path)
    image = tmp_path / "a|b.png"
    _write_png(image)

    result = process_markdown_links(
        "![pipe](a%7Cb.png)",
        source,
        link_config=LinkRuntimeConfig(),
        target_format="xlsx",
        image_scope="posix-scope",
    )

    assert "{{IMAGE@posix-scope:" in result
    assert "%7C" in result
    assert "%257C" not in result
    assert str(image).replace("|", "%7C") in result


@pytest.mark.skipif(os.name == "nt", reason="physical POSIX absolute-path contract")
def test_posix_absolute_image_path_preserves_root(tmp_path: Path) -> None:
    source = _source(tmp_path)
    image = tmp_path / "absolute.png"
    _write_png(image)

    result = process_markdown_links(
        f"![absolute](<{image.as_posix()}>)",
        source,
        link_config=LinkRuntimeConfig(),
        target_format="xlsx",
        image_scope="posix-scope",
    )

    assert "{{IMAGE@posix-scope:" in result
    assert image.as_posix() in result
