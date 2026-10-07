"""Markup source decoding and resource-identity regressions."""

from __future__ import annotations

from ._input_routes_support import (
    Path,
    _document_node_root,
    _run_request,
    _test_png_bytes,
    pipeline as pipeline,
    pytest,
)

pytestmark = pytest.mark.integration


def test_html_honors_admitted_encoding_and_declared_charset(pipeline, tmp_path: Path) -> None:
    _plugin, task_mgr, _ws_mgr = pipeline
    source = tmp_path / "gb.html"
    html = (
        '<html><head><meta charset="gb18030"><title>中文标题</title></head>'
        '<body><p>中文正文</p></body></html>'
    )
    source.write_bytes(html.encode("gb18030"))
    output = tmp_path / "out-html-encoding"
    output.mkdir()

    result = _run_request(
        task_mgr,
        source,
        "html",
        output,
        input_encoding="gb18030",
    )

    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "中文标题" in markdown
    assert "中文正文" in markdown
    assert "\ufffd" not in markdown


def test_html_companion_folder_prefix_is_not_joined_twice(pipeline, tmp_path: Path) -> None:
    _plugin, task_mgr, _ws_mgr = pipeline
    source = tmp_path / "saved.html"
    resources = tmp_path / "saved_files"
    resources.mkdir()
    expected = _test_png_bytes((12, 34, 56))
    (resources / "chart.png").write_bytes(expected)
    source.write_text(
        '<html><head><title>Saved</title></head>'
        '<body><img src="saved_files/chart.png" alt="chart"></body></html>',
        encoding="utf-8",
    )
    output = tmp_path / "out-html-resource"
    output.mkdir()

    result = _run_request(
        task_mgr,
        source,
        "html",
        output,
        image_link_style="markdown_embed",
    )

    assert result.success, result.error
    images = [item for item in result.artifacts if item.kind == "image"]
    assert len(images) == 1
    assert Path(images[0].staging_path).read_bytes() == expected
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "chart.png" in markdown
    assert "saved_files/saved_files" not in markdown


def test_epub_same_basename_images_resolve_relative_to_each_chapter(pipeline, tmp_path: Path) -> None:
    from ebooklib import epub

    _plugin, task_mgr, _ws_mgr = pipeline
    book = epub.EpubBook()
    book.set_identifier("same-basename-images")
    book.set_title("Same Basename")
    book.set_language("en")

    first_bytes = _test_png_bytes((220, 20, 20))
    second_bytes = _test_png_bytes((20, 20, 220))
    book.add_item(
        epub.EpubItem(
            uid="first-image",
            file_name="one/images/pic.png",
            media_type="image/png",
            content=first_bytes,
        )
    )
    book.add_item(
        epub.EpubItem(
            uid="second-image",
            file_name="two/images/pic.png",
            media_type="image/png",
            content=second_bytes,
        )
    )

    first = epub.EpubHtml(title="First", file_name="one/chapter.xhtml", lang="en")
    first.content = '<html><body><p><img src="images/pic.png" alt="first"></p></body></html>'
    second = epub.EpubHtml(title="Second", file_name="two/chapter.xhtml", lang="en")
    second.content = '<html><body><p><img src="images/pic.png" alt="second"></p></body></html>'
    book.add_item(first)
    book.add_item(second)
    book.toc = [first, second]
    book.spine = ["nav", first, second]
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())

    source = tmp_path / "same.epub"
    epub.write_epub(str(source), book)
    output = tmp_path / "out-epub-resource"
    output.mkdir()

    result = _run_request(
        task_mgr,
        source,
        "epub",
        output,
        image_link_style="markdown_embed",
    )

    assert result.success, result.error
    images = [item for item in result.artifacts if item.kind == "image"]
    assert [item.suggested_name for item in images] == ["pic.png", "pic-2.png"]
    assert Path(images[0].staging_path).read_bytes() == first_bytes
    assert Path(images[1].staging_path).read_bytes() == second_bytes
    assert _document_node_root(Path(images[0].staging_path), output) == _document_node_root(
        Path(images[1].staging_path),
        output,
    )

    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    first_link = "![first](pic.png)"
    second_link = "![second](pic-2.png)"
    assert first_link in markdown
    assert second_link in markdown
    assert markdown.index(first_link) < markdown.index(second_link)
