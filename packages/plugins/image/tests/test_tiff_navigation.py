"""TIFF results must be navigable while preserving physical-page Bundle facts."""

from __future__ import annotations

import hashlib
import re
from pathlib import Path
from urllib.parse import unquote, urlsplit

import pytest
from PIL import Image

from docwen_core.models import OutputPolicy
from docwen_core.text.ocr import OcrOutcome, OcrStatus
from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter
from docwen_runtime.output.finalizer import OutputFinalizer

from ._image_conversions_support import _build_fake_context


def _links(text: str) -> list[tuple[bool, str]]:
    return [(bool(embed), target) for embed, target in re.findall(r"(!?)\[[^\]]*\]\(([^)]+)\)", text)]


@pytest.mark.contract
@pytest.mark.parametrize("frames", [1, 4])
@pytest.mark.parametrize("enable_ocr", [False, True])
@pytest.mark.parametrize("keep_images", [False, True])
def test_tiff_navigation_follows_frame_ownership(
    tmp_path: Path,
    sample_four_frame_tiff_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    frames: int,
    enable_ocr: bool,
    keep_images: bool,
) -> None:
    source = tmp_path / "页面 # % (图).tiff"
    with Image.open(sample_four_frame_tiff_path) as original:
        if frames == 1:
            original.save(source, format="TIFF")
        else:
            source.write_bytes(sample_four_frame_tiff_path.read_bytes())
    outcomes = [
        OcrOutcome(OcrStatus.SUCCESS, text="FIRST PAGE CONTENT"),
        OcrOutcome(OcrStatus.NO_TEXT),
        OcrOutcome(OcrStatus.RECOGNITION_FAILED),
        OcrOutcome(OcrStatus.SUCCESS, text="LAST PAGE CONTENT"),
    ]
    calls: list[str] = []

    def ocr(path: str, **_kwargs: object) -> OcrOutcome:
        calls.append(path)
        return outcomes[len(calls) - 1]

    monkeypatch.setattr("docwen_plugin_image.to_markdown.converter.run_ocr_outcome", ocr)
    stage = tmp_path / "staging"
    stage.mkdir()
    result = ImageToMarkdownConverter().convert(
        _build_fake_context(
            str(source),
            str(stage),
            "md",
            {"to_md_enable_ocr": enable_ocr, "to_md_keep_images": keep_images},
            source_format="tif",
        )
    )
    assert result.success
    primary = next(item for item in result.artifacts if item.is_primary)
    fragments = [item for item in result.artifacts if item.kind == "auxiliary"]
    resources = [item for item in result.artifacts if item.kind == "image"]
    assert len(fragments) == (frames if enable_ocr else 0)
    assert len(resources) == (frames if keep_images else 0)
    assert len(calls) == (frames if enable_ocr else 0)
    primary_text = Path(primary.staging_path).read_text(encoding="utf-8")
    targets = _links(primary_text)
    expected = fragments if enable_ocr else resources
    assert [unquote(target) for _, target in targets] == [item.suggested_name for item in expected]
    assert all(embed == (not enable_ocr) for embed, _ in targets)
    assert "FIRST PAGE CONTENT" not in primary_text and "LAST PAGE CONTENT" not in primary_text
    for _, target in targets:
        assert urlsplit(target).fragment == ""  # A literal # in a filename is not an anchor.
        assert " " not in target and "(" not in target and ")" not in target
    for page, fragment in enumerate(fragments, 1):
        text = Path(fragment.staging_path).read_text(encoding="utf-8")
        image_links = _links(text)
        assert [(embed, unquote(target)) for embed, target in image_links] == (
            [(True, resources[page - 1].suggested_name)] if keep_images else []
        )
        assert fragment.metadata["source_page"] == page
        assert fragment.metadata["ocr_status"] == outcomes[page - 1].status.value
        content = outcomes[page - 1].recognized_text
        if content:
            assert text.count(content) == 1
    with Image.open(source) as original:
        for page, resource in enumerate(resources):
            original.seek(page)
            with Image.open(resource.staging_path) as actual:
                assert actual.convert("RGBA").tobytes() == original.convert("RGBA").tobytes()


@pytest.mark.integration
@pytest.mark.parametrize("enable_ocr", [False, True])
def test_tiff_navigation_survives_final_directory_relocation_and_collision(
    tmp_path: Path,
    sample_four_frame_tiff_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    enable_ocr: bool,
) -> None:
    from docwen_application.bundle_mapping import build_bundle_draft
    from docwen_core.models import validate_artifact_bundle_draft
    from docwen_core.models.document_node import ConversionIdentity

    source = tmp_path / "页面 # % (图).tif"
    source.write_bytes(sample_four_frame_tiff_path.read_bytes())
    monkeypatch.setattr(
        "docwen_plugin_image.to_markdown.converter.run_ocr_outcome",
        lambda *_args, **_kwargs: OcrOutcome(OcrStatus.SUCCESS, text="OCR BODY"),
    )
    stage = tmp_path / "staging"
    stage.mkdir()
    converted = ImageToMarkdownConverter().convert(
        _build_fake_context(
            str(source),
            str(stage),
            "md",
            {
                "to_md_enable_ocr": enable_ocr,
                "to_md_keep_images": True,
            },
            source_format="tif",
        )
    )
    assert converted.success
    primary = next(item for item in converted.artifacts if item.is_primary)
    primary_text = Path(primary.staging_path).read_text(encoding="utf-8")
    assert len(_links(primary_text)) == 4
    # Two finalizations of the same source identity collide and must rebase the whole root.
    finalizer = OutputFinalizer()
    output = tmp_path / "output"
    roots: list[str] = []
    identity = ConversionIdentity.create(task_id="tiff.navigation", source_stem=source.stem, source_format="tif")
    for _ in range(2):
        result = finalizer.finalize(
            "tiff.navigation",
            converted.artifacts,
            OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
            input_path=str(source),
            identity=identity,
        )
        assert result.success
        roots.append(result.metrics.extra["document_node_root"])
        by_path = {Path(item.staging_path).resolve(): item for item in result.artifacts}
        actual_primary = next(item for item in result.artifacts if item.is_primary)
        page_targets: list[str] = []
        for artifact in result.artifacts:
            data = Path(artifact.staging_path).read_bytes()
            assert artifact.sha256 == hashlib.sha256(data).hexdigest()
            if artifact.media_type != "text/markdown":
                continue
            text = data.decode("utf-8")
            for embed, target in _links(text):
                target_path = (Path(artifact.staging_path).parent / unquote(target)).resolve()
                linked = by_path[target_path]
                assert linked.media_type == ("image/png" if embed else "text/markdown")
                page_targets.append(linked.artifact_id)
            if artifact.is_primary:
                assert "OCR BODY" not in text
        assert len(page_targets) == 4 * (2 if enable_ocr else 1)
        assert len(set(page_targets)) == len(page_targets)
        draft = build_bundle_draft(
            profile="physical_page_ocr",
            output_media_type="text/markdown",
            artifacts=result.artifacts,
        )
        validate_artifact_bundle_draft(draft)
        assert len(draft.entries) == 1 and draft.entries[0].artifact_id == actual_primary.artifact_id
        pages = [relation for relation in draft.relations if relation.role == "ocr_page"]
        assert len(pages) == (4 if enable_ocr else 0)
        images = [relation for relation in draft.relations if relation.type == "resource_of"]
        assert len(images) == 4
        owners = {relation.source_artifact_id for relation in pages} if enable_ocr else {actual_primary.artifact_id}
        assert all(relation.target_artifact_id in owners for relation in images)
    assert roots[0] != roots[1]


@pytest.mark.contract
def test_cancellation_after_last_frame_prevents_navigation_publication(
    tmp_path: Path,
    sample_four_frame_tiff_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_plugin_image.to_markdown.converter as converter
    from docwen_core.errors import CancellationRequested

    stage = tmp_path / "stage"
    stage.mkdir()
    context = _build_fake_context(
        str(sample_four_frame_tiff_path),
        str(stage),
        "md",
        {
            "to_md_enable_ocr": False,
            "to_md_keep_images": True,
        },
        source_format="tif",
    )
    save = converter.save_image_with_options
    calls = 0

    def cancel_last(image: Image.Image, output_path: str, target_format: str, options: dict[str, object]) -> None:
        nonlocal calls
        save(image, output_path, target_format, options)
        calls += 1
        if calls == 4:
            context._cancellation.cancel("cancel after final frame")

    monkeypatch.setattr(converter, "save_image_with_options", cancel_last)
    with pytest.raises(CancellationRequested):
        ImageToMarkdownConverter().convert(context)
    assert context.workspace.registered_artifacts == []
    assert list(stage.iterdir()) == []
