from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_plugin_layout.to_markdown import converter
from docwen_plugin_layout.to_markdown.converter import (
    _extract_page_markdown,
    _index_extracted_image_targets,
    _referenced_extracted_images,
    _rename_preprocessed_image_refs,
    _rewrite_extracted_image_refs,
)

from ._layout_conversions_support import _build_fake_context

pytestmark = pytest.mark.contract


def _write_png(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with Image.new("RGB", (2, 2), (1, 2, 3)) as image:
        image.save(path, format="PNG")


def _target_index(path: Path, prefix: str) -> dict[str, set[Path]]:
    paths_by_target: dict[str, set[Path]] = {}
    _index_extracted_image_targets(paths_by_target, [path], images_prefix=prefix)
    return paths_by_target


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
        markdown, written_images = _extract_page_markdown(renderer, document, 2, directory)

    (image_file,) = directory.iterdir()
    assert written_images == [image_file]
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
    assert _referenced_extracted_images(wiki, _target_index(image_file, "ignored/")) == {image_file}


def test_layout_indexes_each_written_image_once_and_keeps_cross_page_references(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    renderer = pytest.importorskip("pymupdf4llm")
    pymupdf = pytest.importorskip("pymupdf")
    from pymupdf4llm.helpers import document_layout

    page_count = 6
    pdf = tmp_path / "source.pdf"
    with pymupdf.open() as document:
        for _ in range(page_count):
            document.new_page()
        document.save(pdf)

    picture = tmp_path / "picture.png"
    _write_png(picture)
    picture_bytes = picture.read_bytes()

    def parse_document(_document, *, pages, **_):
        (index,) = pages
        box = SimpleNamespace(image=picture_bytes)

        def to_markdown(**_):
            extra = "\n![](<source.pdf-0001-00.png>)" if index == 1 else ""
            return f"Page {index + 1}\n![]({box.image}){extra}"

        return SimpleNamespace(pages=[SimpleNamespace(boxes=[box])], to_markdown=to_markdown)

    detected = converter._detected_image_format
    classifications: list[str] = []

    def count_detection(path: Path) -> str | None:
        classifications.append(path.name)
        return detected(path)

    monkeypatch.setattr(renderer, "_use_layout", True)
    monkeypatch.setattr(document_layout, "parse_document", parse_document)
    monkeypatch.setattr(converter, "_detected_image_format", count_detection)
    result = converter.LayoutToMarkdownConverter().convert(
        _build_fake_context(str(pdf), str(tmp_path), "md", options={"to_md_keep_images": True})
    )

    assert result.success, result.error
    images = [artifact for artifact in result.artifacts if artifact.kind == "image"]
    assert [artifact.metadata.get("source_page") for artifact in images] == [None, 2, 3, 4, 5, 6]
    assert all(Path(artifact.staging_path).read_bytes() == picture_bytes for artifact in images)
    assert [diagnostic.code for diagnostic in result.diagnostics].count("resource_page_unresolved") == 1
    assert sorted(classifications) == sorted(artifact.suggested_name for artifact in images for _ in range(2))


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
    assert _referenced_extracted_images(wiki, _target_index(renamed_image, "ignored/")) == {renamed_image}


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
        f"![]({canonical_target})", _target_index(relative_image, "document_images/")
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
