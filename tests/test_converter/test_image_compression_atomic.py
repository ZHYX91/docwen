"""converter 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from docwen.converter.formats.image.compression import compress_to_size

pytestmark = pytest.mark.unit


def test_compress_to_size_does_not_leave_tmp_file(tmp_path: Path) -> None:
    img = Image.new("RGB", (10, 10), (255, 255, 255))

    out = tmp_path / "out.jpg"
    ok = compress_to_size(img, str(out), "jpg", target_size=1024, unit="KB")

    assert ok is True
    assert out.exists() is True
    assert (tmp_path / "out.jpg.tmp").exists() is False
