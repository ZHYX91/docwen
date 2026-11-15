"""Embedded HTML must not redirect an authored Markdown document."""

from pathlib import Path

import pytest

from docwen_core.detection import detect_content_format, inspect_file

pytestmark = pytest.mark.contract


REFERENCES = """# Reference markers

![full][first]

![collapsed][]

![shortcut]

[first]: assets/green-marker.png
[collapsed]: assets/green-marker.png
[shortcut]: assets/green-marker.png

<!-- ## Hidden HTML heading -->

<div>
## Block HTML heading
</div>

## Visible heading

```text
## Code heading
```
"""


@pytest.mark.parametrize("suffix", [".md", ".txt", ".html"])
@pytest.mark.parametrize(
    "content",
    [
        REFERENCES,
        "# Heading\n\nA paragraph with <div>inline HTML</div>.\n",
        "<div>HTML block</div>\n\n# Visible Markdown heading\n",
        "```html\n<!DOCTYPE html>\n<html><body>Example</body></html>\n```\n",
        "~~~html\n<!DOCTYPE html>\n<html><body>Example</body></html>\n~~~\n",
        "# Heading\n\n<!-- <html> is only a comment -->\n",
        "# Real Markdown\n\n    <html><body>example</body></html>\n",
        "# Real Markdown\n\n\t<html><body>example</body></html>\n",
        "<!-- closed -->" * 100 + "<!-- <html> is not an envelope -->\n\n# Visible Markdown\n",
    ],
)
def test_markdown_with_html_is_detected_independently_of_suffix(tmp_path: Path, suffix: str, content: str) -> None:
    source = tmp_path / ("mixed" + suffix)
    source.write_text(content, encoding="utf-8")
    assert detect_content_format(str(source)).format == "markdown"
    if suffix == ".md":
        assert inspect_file(str(source)).decision == "allow"
    elif suffix == ".html":
        assert inspect_file(str(source)).decision == "require_explicit_acceptance"


@pytest.mark.parametrize(
    "content",
    [
        "<!DOCTYPE html>\n<html>\n<body>\n\n# Not a Markdown heading\n\n</body></html>\n",
        "<!-- comment -->\n<html><body>![text](not-an-image)</body></html>\n",
        "<!-- lead --><html><body>\n\n# HTML visible text\n\n</body></html>\n",
        '<?xml version="1.0"?><html><body>\n\n# HTML visible text\n\n</body></html>\n',
        "<!-- lead --><?xml version='1.0'?><!DOCTYPE html>\n\n# HTML text\n",
        "<div>\n# HTML content\n![text](image.png)\n</div>\n",
        "<p>HTML paragraph</p>\n",
    ],
)
def test_actual_html_still_requires_acceptance_when_named_markdown(tmp_path: Path, content: str) -> None:
    source = tmp_path / "wrong.md"
    source.write_text(content, encoding="utf-8")
    assert detect_content_format(str(source)).format == "html"
    assert inspect_file(str(source)).decision == "require_explicit_acceptance"
