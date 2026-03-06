"""converter 单元测试。"""

from __future__ import annotations

import os
from pathlib import Path

import pytest
from PIL import Image

from docwen.converter.formats.image.core import convert_image

pytestmark = pytest.mark.unit


def _write_noise_bmp(path: Path, size: tuple[int, int] = (512, 512)) -> None:
    w, h = size
    img = Image.frombytes("RGB", (w, h), os.urandom(w * h * 3))
    img.save(path, format="BMP")


def test_limit_size_accepts_lowercase_unit_and_meets_target(tmp_path: Path) -> None:
    src = tmp_path / "in.bmp"
    _write_noise_bmp(src)

    out = tmp_path / "out.jpg"
    convert_image(
        str(src),
        "jpg",
        str(out),
        {"compress_mode": "limit_size", "size_limit": 500, "size_unit": "kb"},
    )

    assert out.exists() is True
    assert out.stat().st_size <= 500 * 1024


def test_limit_size_impossible_target_raises(tmp_path: Path) -> None:
    src = tmp_path / "in.bmp"
    _write_noise_bmp(src)

    out = tmp_path / "out.webp"
    with pytest.raises(RuntimeError) as exc:
        convert_image(
            str(src),
            "webp",
            str(out),
            {"compress_mode": "limit_size", "size_limit": 1, "size_unit": "KB"},
        )

    assert "无法将图片压缩到目标大小" in str(exc.value)


def test_limit_size_invalid_unit_raises(tmp_path: Path) -> None:
    src = tmp_path / "in.bmp"
    _write_noise_bmp(src, (64, 64))

    out = tmp_path / "out.jpg"
    with pytest.raises(RuntimeError) as exc:
        convert_image(
            str(src),
            "jpg",
            str(out),
            {"compress_mode": "limit_size", "size_limit": 10, "size_unit": "KiB"},
        )

    assert "无效的单位" in str(exc.value)


def test_limit_size_does_not_skip_when_source_is_small(tmp_path: Path) -> None:
    w, h = 256, 256
    img = Image.frombytes("RGB", (w, h), os.urandom(w * h * 3))
    src = tmp_path / "in.webp"
    img.save(src, format="WEBP", quality=10, method=6)

    baseline = tmp_path / "baseline.jpg"
    convert_image(
        str(src),
        "jpg",
        str(baseline),
        {"compress_mode": "lossless"},
    )

    src_size = src.stat().st_size
    baseline_size = baseline.stat().st_size
    assert src_size > 0
    assert baseline_size > 0
    assert src_size < baseline_size

    target_bytes = (src_size + baseline_size) // 2
    target_kb = max((target_bytes + 1023) // 1024, (src_size + 1023) // 1024 + 1)

    out = tmp_path / "out.jpg"
    convert_image(
        str(src),
        "jpg",
        str(out),
        {"compress_mode": "limit_size", "size_limit": int(target_kb), "size_unit": "KB"},
    )

    assert out.exists() is True
    assert out.stat().st_size <= int(target_kb) * 1024
