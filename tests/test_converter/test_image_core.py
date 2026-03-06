"""converter.formats.image.core 的单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from docwen.converter.formats.image.core import convert_image

pytestmark = pytest.mark.unit


@pytest.mark.slow
def test_convert_image_rgba_to_jpeg(tmp_path: Path) -> None:
    src = tmp_path / "in.png"
    Image.new("RGBA", (2, 2), (255, 0, 0, 128)).save(src)

    out = tmp_path / "out.jpg"
    convert_image(str(src), "jpeg", str(out), {"compress_mode": "lossless"})
    assert out.exists() is True

    with Image.open(out) as img:
        assert img.mode == "RGB"
