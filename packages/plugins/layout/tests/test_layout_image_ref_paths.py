from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_plugin_layout.to_markdown.converter import (
    _extract_page_markdown,
    _referenced_extracted_images,
    _rename_preprocessed_image_refs,
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
    pymupdf = pytest.importorskip("pymupdf")
    from pymupdf4llm.helpers import document_layout

    directory = tmp_path / "folder with spaces [and brackets]"
    directory.mkdir()
    source = tmp_path / "original.png"
    _write_png(source)
    box = SimpleNamespace(image=source.read_bytes())
    page = SimpleNamespace(boxes=[box])
    parsed = SimpleNamespace(pages=[page], to_markdown=lambda **kwargs: f"![]({box.image})")
    monkeypatch.setattr(renderer, "_use_layout", True)
    monkeypatch.setattr(document_layout, "parse_document", lambda *args, **kwargs: parsed)

    pdf = tmp_path / "report [draft] #2.pdf"
    with pymupdf.open() as document:
        for _ in range(3):
            document.new_page()
        document.save(pdf)
    with pymupdf.open(pdf) as document:
        markdown = _extract_page_markdown(renderer, document, 2, directory)

    (image_file,) = directory.iterdir()
    assert image_file.name == "report draft 2.pdf-0003-00.png"
    assert image_file.read_bytes() == source.read_bytes()
    assert isinstance(markdown, str)
    assert markdown == f"![](<{image_file.name}>)"

    wiki, _ = _rewrite_extracted_image_refs(
        markdown,
        images_prefix="ignored/",
        image_files=[image_file],
        image_mode="file",
        image_link_style="wiki_embed",
    )
    assert wiki == f"![[{image_file.name}]]"
    assert _referenced_extracted_images(wiki, images_prefix="ignored/", image_files={image_file}) == {image_file}


def test_preprocessed_ofd_renames_image_and_preserves_resolvable_wiki_targets(tmp_path: Path) -> None:
    effective_input = tmp_path / "_preprocess_ofd_cafe0001.pdf"
    original_input = tmp_path / "report [draft] #2.ofd"
    image_file = tmp_path / f"{effective_input.name}-0001-00.png"
    _write_png(image_file)
    payload = image_file.read_bytes()
    markdown = f"![](<{image_file.name}>)\n![](<{image_file.as_posix()}>)"

    markdown, images = _rename_preprocessed_image_refs(
        markdown,
        image_files=[image_file],
        effective_input_path=str(effective_input),
        original_input_path=str(original_input),
    )
    (renamed_image,) = images
    assert renamed_image.name == "report draft 2.ofd-0001-00.png"
    assert renamed_image.read_bytes() == payload
    assert not image_file.exists()

    wiki, _ = _rewrite_extracted_image_refs(
        markdown,
        images_prefix="ignored/",
        image_files=images,
        image_mode="file",
        image_link_style="wiki_embed",
    )
    assert wiki.splitlines() == [f"![[{renamed_image.name}]]"] * 2
    assert _referenced_extracted_images(wiki, images_prefix="ignored/", image_files={renamed_image}) == {renamed_image}


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
