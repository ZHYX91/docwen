"""Final HTML-to-Markdown text keeps authored spellings of former image tokens."""

from pathlib import Path

import pytest

from docwen_plugin_layout.to_markdown.converter import convert_html_to_markdown_text

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("keep", [False, True])
def test_final_layout_html_consumer_keeps_body_and_code_tokens(tmp_path: Path, keep: bool) -> None:
    output = tmp_path / "output"
    output.mkdir()
    content = convert_html_to_markdown_text(
        html_text='<p>DOCWENIMG0001</p><pre><code>DOCWENIMG0001</code></pre><img src="https://example.invalid/image.png"/>',
        html_path=str(tmp_path / "page.html"),
        output_folder=str(output),
        keep_images=keep,
    )
    assert content.count("DOCWENIMG0001") == 2
