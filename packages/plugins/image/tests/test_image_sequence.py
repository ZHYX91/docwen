"""Persisted sequence semantics replace historical first-frame projections."""

from pathlib import Path

import pytest
from PIL import Image

from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

from ._image_conversions_support import _build_fake_context

pytestmark = [pytest.mark.golden, pytest.mark.contract]


@pytest.mark.parametrize("target", ["gif", "webp", "png", "tif", "jpg", "bmp"])
def test_animation_retains_all_frames(tmp_path: Path, target: str) -> None:
    source = tmp_path / "animation.gif"
    frames = [Image.new("RGB", (20, 20), color) for color in ("red", "green", "blue")]
    try:
        frames[0].save(source, save_all=True, append_images=frames[1:], duration=[40, 80, 120], loop=2)
    finally:
        for frame in frames:
            frame.close()
    staging = tmp_path / "staging"
    staging.mkdir()
    result = ImageFormatConverter().convert(_build_fake_context(str(source), str(staging), target))
    assert result.success
    assert result.metrics.extra["frame_count"] == 3
    split = target in ("jpg", "bmp")
    assert len(result.artifacts) == (3 if split else 1)
    with Image.open(result.artifacts[0].staging_path) as output:
        assert getattr(output, "n_frames", 1) == (1 if split else 3)
        if target in ("gif", "png", "webp"):
            durations = []
            for index in range(3):
                output.seek(index)
                output.load()
                durations.append(output.info["duration"])
            assert durations == [40, 80, 120]
            assert output.info["loop"] == (2 if target == "gif" else 3)
    warnings = [diagnostic.code for diagnostic in result.diagnostics if diagnostic.level == "warning"]
    assert bool(warnings) == (target in ("tif", "jpg", "bmp"))


def test_multiframe_tiff_retains_pages(tmp_path: Path) -> None:
    source = tmp_path / "pages.tif"
    with Image.new("RGB", (20, 20), "red") as first, Image.new("RGB", (10, 10), "blue") as second:
        first.save(source, save_all=True, append_images=[second])
    result = ImageFormatConverter().convert(_build_fake_context(str(source), str(tmp_path), "tif"))
    assert result.success
    with Image.open(result.artifacts[0].staging_path) as output:
        assert getattr(output, "n_frames", 1) == 2
        assert output.size == (20, 20)
        output.seek(1)
        assert output.size == (10, 10)


@pytest.mark.parametrize("plays", [0, 1, 3])
def test_apng_to_gif_preserves_total_play_count(tmp_path: Path, plays: int) -> None:
    source = tmp_path / "animation.png"
    with Image.new("RGBA", (20, 20), "red") as first, Image.new("RGBA", (20, 20), "blue") as second:
        first.save(source, save_all=True, append_images=[second], duration=100, loop=plays)
    result = ImageFormatConverter().convert(_build_fake_context(str(source), str(tmp_path), "gif"))
    assert result.success
    with Image.open(result.artifacts[0].staging_path) as output:
        assert getattr(output, "n_frames", 1) == 2
        assert output.info.get("loop") == ({0: 0, 1: None, 3: 2}[plays])


def test_nonlooping_gif_does_not_become_infinite_webp(tmp_path: Path) -> None:
    source = tmp_path / "once.gif"
    with Image.new("RGB", (20, 20), "red") as first, Image.new("RGB", (20, 20), "blue") as second:
        first.save(source, save_all=True, append_images=[second], duration=100)
    result = ImageFormatConverter().convert(_build_fake_context(str(source), str(tmp_path), "webp"))
    assert result.success
    with Image.open(result.artifacts[0].staging_path) as output:
        assert output.info["loop"] == 1


def test_animation_size_limit_does_not_flatten_sequence(tmp_path: Path) -> None:
    source = tmp_path / "animation.gif"
    with Image.new("RGB", (20, 20), "red") as first, Image.new("RGB", (20, 20), "blue") as second:
        first.save(source, save_all=True, append_images=[second], duration=100, loop=0)
    result = ImageFormatConverter().convert(
        _build_fake_context(
            str(source), str(tmp_path), "webp", {"compress_mode": "limit_size", "size_limit": 10, "size_unit": "KB"}
        )
    )
    assert result.success
    with Image.open(result.artifacts[0].staging_path) as output:
        assert getattr(output, "n_frames", 1) == 2
    assert Path(result.artifacts[0].staging_path).stat().st_size <= 10 * 1024
