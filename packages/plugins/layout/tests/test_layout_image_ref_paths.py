from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_plugin_layout.to_markdown.converter import (
    _referenced_extracted_images,
    _rewrite_extracted_image_refs,
)

pytestmark = pytest.mark.contract


def _write_png(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with Image.new("RGB", (2, 2), (1, 2, 3)) as image:
        image.save(path, format="PNG")


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
