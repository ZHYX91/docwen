"""Saved image content, bounded merge work and explicit WebP lossless encoding."""

from __future__ import annotations

import pytest
from PIL import Image

from docwen_core.errors import CancellationRequested
from docwen_plugin_image._common import save_image_with_options
from docwen_plugin_image.merge import converter as merge

from ._image_conversions_support import _build_fake_context

pytestmark = pytest.mark.contract


def _merge_context(tmp_path, options=None):
    source = tmp_path / "pages.tif"
    frames = [Image.new("RGBA", (2, 2), color) for color in [(255, 0, 0, 0), (0, 255, 0, 96), (0, 0, 255, 255)]]
    try:
        frames[0].save(source, save_all=True, append_images=frames[1:])
    finally:
        for frame in frames:
            frame.close()
    extra = tmp_path / "opaque.png"
    with Image.new("RGB", (2, 2), (10, 20, 30)) as opaque:
        opaque.save(extra)
    staging = tmp_path / "staging"
    staging.mkdir()
    return _build_fake_context(
        str(source),
        str(staging),
        "tif",
        options or {"keep_alpha": True},
        action_name="merge_images_to_tiff",
        extra_input_paths=[str(extra)],
    )


@pytest.mark.parametrize("options", [{"mode": "smart"}, {"keep_alpha": True}])
def test_merge_preserves_each_tiff_page_and_mixed_input_alpha(tmp_path, options):
    result = merge.ImageToTiffMerger().convert(_merge_context(tmp_path, options))
    assert result.success
    assert result.artifacts[0].metadata["image_count"] == 2
    assert result.artifacts[0].metadata["frame_count"] == 4
    with Image.open(result.artifacts[0].staging_path) as saved:
        assert getattr(saved, "n_frames", 1) == 4
        for index, color in enumerate([(255, 0, 0, 0), (0, 255, 0, 96), (0, 0, 255, 255), (10, 20, 30, 255)]):
            saved.seek(index)
            assert saved.mode == "RGBA"
            assert saved.getpixel((0, 0)) == color


@pytest.mark.parametrize("options", [{"mode": "rgb"}, {"keep_alpha": False}])
def test_merge_explicit_rgb_still_flattens_transparency(tmp_path, options):
    result = merge.ImageToTiffMerger().convert(_merge_context(tmp_path, options))
    assert result.success
    with Image.open(result.artifacts[0].staging_path) as saved:
        assert getattr(saved, "n_frames", 1) == 4
        assert saved.mode == "RGB"
        assert saved.getpixel((0, 0)) == (255, 255, 255)


@pytest.mark.parametrize(("budget", "limit"), [("MAX_MERGE_FRAMES", 2), ("MAX_MERGE_PIXELS", 8)])
def test_merge_budget_failure_does_not_publish_partial_pages(tmp_path, monkeypatch, budget, limit):
    monkeypatch.setattr(merge, budget, limit)
    result = merge.ImageToTiffMerger().convert(_merge_context(tmp_path))
    assert not result.success
    assert result.error is not None
    assert result.error.diagnostic_code == "IMG2TIFF-LIMIT"
    assert "Merge smaller groups" in result.error.message
    assert result.artifacts == []
    assert not list((tmp_path / "staging").iterdir())


def test_merge_cancellation_between_frames_is_propagated(tmp_path, monkeypatch):
    context = _merge_context(tmp_path)
    checks = 0

    def cancel_mid_input():
        nonlocal checks
        checks += 1
        if checks == 4:
            raise CancellationRequested("cancel between TIFF pages")

    monkeypatch.setattr(context.cancellation, "check", cancel_mid_input)
    with pytest.raises(CancellationRequested):
        merge.ImageToTiffMerger().convert(context)
    assert not list((tmp_path / "staging").iterdir())


@pytest.mark.parametrize("mode", ["RGB", "RGBA"])
def test_webp_lossless_preserves_saved_pixels_including_hidden_rgb(tmp_path, mode):
    source = Image.new(mode, (32, 32))
    for y in range(32):
        for x in range(32):
            color = ((x * 37 + y * 11) % 256, (x * 13 + y * 53) % 256, (x * 73 + y * 7) % 256)
            source.putpixel((x, y), (*color, (x * 17) % 256) if mode == "RGBA" else color)
    destination = tmp_path / "result.webp"
    try:
        save_image_with_options(source, str(destination), "webp", {"compress_mode": "lossless"})
        with Image.open(destination) as saved:
            assert saved.mode == mode
            assert saved.tobytes() == source.tobytes()
    finally:
        source.close()


def test_legacy_lossless_option_remains_supported_for_jpeg(tmp_path):
    with Image.new("RGB", (2, 2), (20, 40, 60)) as source:
        destination = tmp_path / "result.jpg"
        save_image_with_options(source, str(destination), "jpg", {"compress_mode": "lossless"})
    with Image.open(destination) as saved:
        assert saved.format == "JPEG"
        assert saved.size == (2, 2)
