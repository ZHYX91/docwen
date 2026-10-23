"""EPUB ZIP member identity must remain case-sensitive through Markdown links."""

from __future__ import annotations

from pathlib import Path

import pytest

from ._input_routes_support import _run_request, _test_png_bytes
from ._input_routes_support import pipeline as pipeline

pytestmark = pytest.mark.integration


def test_epub_image_names_differing_only_in_case_keep_distinct_references(pipeline, tmp_path: Path) -> None:
    from ebooklib import epub

    book = epub.EpubBook()
    book.set_identifier("case-sensitive-image-members")
    book.set_title("Case-sensitive images")
    book.set_language("en")
    upper = _test_png_bytes((220, 20, 20))
    lower = _test_png_bytes((20, 20, 220))
    book.add_item(epub.EpubItem(
        uid="upper-image", file_name="Images/A.png",
        media_type="image/png", content=upper,
    ))
    book.add_item(epub.EpubItem(
        uid="lower-image", file_name="Images/a.png",
        media_type="image/png", content=lower,
    ))
    chapter = epub.EpubHtml(title="Images", file_name="chapters/chapter.xhtml", lang="en")
    chapter.content = (
        "<html><body><p>Different image members</p>"
        "<img src='../Images/A.png' alt='upper'/>"
        "<img src='../Images/a.png' alt='lower'/>"
        "</body></html>"
    )
    book.add_item(chapter)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = ["nav", chapter]
    book.toc = [chapter]

    source = tmp_path / "case-sensitive.epub"
    epub.write_epub(str(source), book)
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    _plugin, task_mgr, _ws_mgr = pipeline
    result = _run_request(task_mgr, source, "epub", output_dir, image_link_style="markdown_embed")

    assert result.success, result.error
    images = [artifact for artifact in result.artifacts if artifact.kind == "image"]
    assert len(images) == 2
    assert Path(images[0].staging_path).read_bytes() == upper
    assert Path(images[1].staging_path).read_bytes() == lower
    assert images[0].suggested_name != images[1].suggested_name
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert f"![upper]({images[0].suggested_name})" in content
    assert f"![lower]({images[1].suggested_name})" in content


@pytest.mark.contract
def test_epub_key_normalization_preserves_case_without_changing_other_routes() -> None:
    from docwen_plugin_markup.markdown_resources import normalize_resource_key

    assert normalize_resource_key("../Images/A.png", case_sensitive=True) == "Images/A.png"
    assert normalize_resource_key("../Images/a.png", case_sensitive=True) == "Images/a.png"
    assert normalize_resource_key("../Images/A.png") == "images/a.png"
