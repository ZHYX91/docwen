"""File round trips for colour-key, palette and grayscale transparency."""

from pathlib import Path

import pytest
from PIL import Image

from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

from ._image_conversions_support import _build_fake_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("mode", ["RGB", "L", "P", "LA"])
@pytest.mark.parametrize("target", ["png", "webp", "tif", "gif", "jpg", "bmp"])
def test_export_reopens_with_transparent_or_white_background(tmp_path: Path, mode: str, target: str) -> None:
    source = tmp_path / "source.png"
    transparent, opaque = {
        "RGB": ((10, 40, 60), (180, 180, 180)),
        "L": (32, 180),
        "P": (0, 1),
        "LA": ((32, 0), (180, 255)),
    }[mode]
    with Image.new(mode, (32, 16), transparent) as image:
        if mode == "P":
            image.putpalette([10, 40, 60, 180, 180, 180])
        image.paste(opaque, (16, 0, 32, 16))
        if mode != "LA":
            image.info["transparency"] = transparent
        image.save(source)
    result = ImageFormatConverter().convert(_build_fake_context(str(source), str(tmp_path), target))
    assert result.success, result.error
    with Image.open(result.artifacts[0].staging_path) as encoded, encoded.convert("RGBA") as decoded:
        pixel = decoded.getpixel((4, 8))
        assert isinstance(pixel, tuple)
        if target in ("jpg", "bmp"):
            assert pixel == pytest.approx((255, 255, 255, 255), abs=2)
        else:
            assert pixel[3] == 0
        assert decoded.getpixel((26, 8)) == pytest.approx((180, 180, 180, 255), abs=2)


def test_gif_reserves_one_index_for_different_fully_transparent_colours(tmp_path: Path) -> None:
    source = tmp_path / "source.png"
    with Image.new("RGBA", (3, 1)) as image:
        image.putdata([(20, 0, 0, 0), (0, 0, 80, 0), (20, 0, 0, 255)])
        image.save(source)
    result = ImageFormatConverter().convert(_build_fake_context(str(source), str(tmp_path), "gif"))
    assert result.success
    with (
        Image.open(result.artifacts[0].staging_path) as encoded,
        encoded.convert("RGBA") as decoded,
        decoded.getchannel("A") as alpha,
    ):
        assert alpha.tobytes() == bytes([0, 0, 255])


@pytest.mark.parametrize("target", ["gif", "tif"])
def test_transparent_regions_survive_on_every_frame(tmp_path: Path, target: str) -> None:
    source = tmp_path / "sequence.png"
    with Image.new("RGBA", (3, 1)) as first, Image.new("RGBA", (3, 1)) as second:
        first.putdata([(25, 0, 0, 0), (0, 0, 80, 0), (20, 0, 0, 255)])
        second.putdata([(0, 0, 80, 0), (60, 0, 0, 255), (20, 0, 0, 0)])
        first.save(source, save_all=True, append_images=[second], duration=[50, 80], blend=0)
    result = ImageFormatConverter().convert(_build_fake_context(str(source), str(tmp_path), target))
    assert result.success, result.error
    with Image.open(result.artifacts[0].staging_path) as output:
        for index, expected in enumerate(([0, 0, 255], [0, 255, 0])):
            output.seek(index)
            with output.convert("RGBA") as decoded, decoded.getchannel("A") as alpha:
                assert alpha.tobytes() == bytes(expected)
