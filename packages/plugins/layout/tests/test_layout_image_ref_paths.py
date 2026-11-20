from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_plugin_layout.to_markdown.converter import (
    _extract_page_markdown,
    _referenced_extracted_images,
    _rewrite_extracted_image_refs,
)

pytestmark = pytest.mark.contract


def _write_png(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with Image.new("RGB", (2, 2), (1, 2, 3)) as image:
        image.save(path, format="PNG")


def test_layout_picture_bytes_keep_physical_path_independent_of_markdown(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    renderer = pytest.importorskip("pymupdf4llm")
    from pymupdf4llm.helpers import document_layout

    directory = tmp_path / "folder with spaces [and brackets]"
    directory.mkdir()
    source = tmp_path / "original.png"
    _write_png(source)
    box = SimpleNamespace(image=source.read_bytes())
    page = SimpleNamespace(boxes=[box])
    parsed = SimpleNamespace(pages=[page], to_markdown=lambda: f"![]({box.image})")
    monkeypatch.setattr(renderer, "_use_layout", True)
    monkeypatch.setattr(document_layout, "parse_document", lambda *args, **kwargs: parsed)

    markdown = _extract_page_markdown(renderer, SimpleNamespace(name="input file.pdf"), 2, directory)
    (image_file,) = directory.iterdir()
    assert image_file.read_bytes() == source.read_bytes()
    assert markdown == f"![](<{image_file.name}>)"


def test_bare_generated_image_name_is_rewritten_to_the_requested_link_style(tmp_path: Path) -> None:
    image_path = tmp_path / "page-01.png"
    _write_png(image_path)

    markdown, _ = _rewrite_extracted_image_refs(
        "![](page-01.png)",
        images_prefix="document_images/",
        image_files=[image_path],
        image_mode="file",
        image_link_style="wiki_embed",
    )

    assert markdown == "![[page-01.png]]"


def test_canonical_image_path_matches_an_equivalent_noncanonical_spelling(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    image_path = tmp_path / "page-01.png"
    _write_png(image_path)
    monkeypatch.chdir(tmp_path.parent)
    relative_image = Path(tmp_path.name) / image_path.name
    canonical_target = image_path.resolve().as_posix()

    assert _referenced_extracted_images(
        f"![]({canonical_target})",
        images_prefix="document_images/",
        image_files={relative_image},
    ) == {relative_image}

    markdown, _ = _rewrite_extracted_image_refs(
        f"![]({canonical_target})",
        images_prefix="document_images/",
        image_files=[relative_image],
        image_mode="base64",
        image_link_style="markdown_embed",
        export_semantics=MarkdownExportSemantics(),
    )
    assert markdown.startswith("![page-01](data:image/png;base64,")
