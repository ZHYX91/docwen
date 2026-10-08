"""Cancellation in recursively expanded Markdown must not become document text."""

from pathlib import Path

import pytest

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_core.links._embed_md import process_embedded_md_file

pytestmark = pytest.mark.unit


def test_embedded_markdown_propagates_cancellation(tmp_path: Path):
    source = tmp_path / "child.md"
    original = b"Child paragraph.\n"
    source.write_bytes(original)
    token = CancellationToken()

    def cancel_child(*args, **kwargs):
        token.cancel("user_cancelled")
        token.check()
        return "unreachable"

    with pytest.raises(CancellationRequested):
        process_embedded_md_file(str(source), str(tmp_path / "parent.md"), None, 0, cancel_child)
    assert source.read_bytes() == original


@pytest.mark.parametrize("kind", ["wiki-image", "markdown-image", "wiki-link", "markdown-link", "bare-url"])
def test_link_batches_stop_before_next_resolution(tmp_path, monkeypatch, kind):
    from docwen_core.export_semantics import LinkRuntimeConfig
    from docwen_core.links import _markdown_orchestrator, _non_embed, process_markdown_links

    token = CancellationToken()
    calls = []

    def resolve(*args, **kwargs):
        calls.append((args, kwargs))
        token.cancel("user_cancelled")
        return "resolved"

    if kind == "wiki-image":
        monkeypatch.setattr(_markdown_orchestrator, "resolve_embedded_links", resolve)
        text = "![[one.png]] ![[two.png]]"
    elif kind == "markdown-image":
        monkeypatch.setattr(_markdown_orchestrator, "process_single_embed", resolve)
        text = "![one](one.png) ![two](two.png)"
    elif kind == "wiki-link":
        monkeypatch.setattr(_non_embed, "resolve_file_path", resolve)
        text = "[[one.md]] [[two.md]]"
    elif kind == "markdown-link":
        monkeypatch.setattr(_non_embed, "_canonical_local_docx_target", resolve)
        text = "[one](one.md) [two](two.md)"
    else:
        bare_url_end = _markdown_orchestrator._bare_url_end

        def cancel_after_url(segment, start):
            end = bare_url_end(segment, start)
            if end is not None:
                calls.append(end)
                token.cancel("user_cancelled")
            return end

        monkeypatch.setattr(_markdown_orchestrator, "_bare_url_end", cancel_after_url)
        text = "https://example.org/one https://example.org/two"
    with pytest.raises(CancellationRequested):
        process_markdown_links(
            text,
            str(tmp_path / "source.md"),
            link_config=LinkRuntimeConfig(auto_link_bare_url=kind == "bare-url"),
            target_format="docx",
            _canonicalize_local_docx_targets=True,
            cancellation_check=token.check,
        )
    assert len(calls) == 1
