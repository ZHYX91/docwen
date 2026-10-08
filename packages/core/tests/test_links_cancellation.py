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
