"""Shared fixtures for image plugin tests."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from PIL import Image

PROJECT_ROOT = Path(__file__).resolve().parents[4]

LOCAL_SRC_PATHS = [
    PROJECT_ROOT,
    PROJECT_ROOT / "packages" / "core" / "src",
    PROJECT_ROOT / "packages" / "runtime" / "src",
    PROJECT_ROOT / "packages" / "plugins" / "image" / "src",
]

for path in reversed(LOCAL_SRC_PATHS):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))


@pytest.fixture
def sample_png_path(tmp_path: Path) -> Path:
    path = tmp_path / "sample.png"
    img = Image.new("RGBA", (32, 24), (20, 120, 200, 128))
    img.save(path, format="PNG")
    img.close()
    return path


@pytest.fixture
def sample_second_png_path(tmp_path: Path) -> Path:
    path = tmp_path / "second.png"
    img = Image.new("RGB", (16, 16), (240, 220, 40))
    img.save(path, format="PNG")
    img.close()
    return path


@pytest.fixture
def sample_tiff_path(tmp_path: Path) -> Path:
    path = tmp_path / "multi.tif"
    first = Image.new("RGB", (12, 10), (255, 0, 0))
    second = Image.new("RGB", (12, 10), (0, 255, 0))
    first.save(path, format="TIFF", save_all=True, append_images=[second], compression="tiff_lzw")
    first.close()
    second.close()
    return path


@pytest.fixture
def sample_four_frame_tiff_path(tmp_path: Path) -> Path:
    path = tmp_path / "four-pages.tif"
    frames = [Image.new("RGB", (12, 10), color) for color in ((255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0))]
    try:
        frames[0].save(path, format="TIFF", save_all=True, append_images=frames[1:], compression="tiff_lzw")
    finally:
        for frame in frames:
            frame.close()
    return path


@pytest.fixture
def sample_jpg_path(tmp_path: Path) -> Path:
    path = tmp_path / "sample.jpg"
    img = Image.new("RGB", (48, 36), (200, 80, 40))
    img.save(path, format="JPEG", quality=90)
    img.close()
    return path


@pytest.fixture
def sample_gif_path(tmp_path: Path) -> Path:
    path = tmp_path / "sample.gif"
    img = Image.new("P", (20, 20))
    img.putpalette([0, 0, 0, 255, 0, 0, 0, 255, 0, 0, 0, 255] + [0, 0, 0] * 252)
    for y in range(3):
        for x in range(3):
            img.putpixel((x + 1, y + 1), y * 3 + x + 1)
    img.save(path, format="GIF")
    img.close()
    return path


@pytest.fixture
def sample_bmp_path(tmp_path: Path) -> Path:
    path = tmp_path / "sample.bmp"
    img = Image.new("RGB", (24, 18), (60, 180, 75))
    img.save(path, format="BMP")
    img.close()
    return path


@pytest.fixture
def sample_webp_path(tmp_path: Path) -> Path:
    path = tmp_path / "sample.webp"
    img = Image.new("RGBA", (40, 30), (128, 64, 192, 200))
    img.save(path, format="WEBP")
    img.close()
    return path


@pytest.fixture
def sample_heic_path(tmp_path: Path) -> Path:
    """Fake HEIC file — creates a dummy file with .heic extension."""
    path = tmp_path / "sample.heic"
    path.write_bytes(b"fake-heic-content-not-a-valid-image")
    return path


@pytest.fixture
def real_heic_path(tmp_path: Path) -> Path:
    pillow_heif = pytest.importorskip("pillow_heif")
    pillow_heif.register_heif_opener()
    path = tmp_path / "sample.heic"
    img = Image.new("RGB", (18, 12), (90, 40, 180))
    img.save(path, format="HEIF")
    img.close()
    return path


@pytest.fixture
def real_heif_path(tmp_path: Path) -> Path:
    pillow_heif = pytest.importorskip("pillow_heif")
    pillow_heif.register_heif_opener()
    path = tmp_path / "sample.heif"
    img = Image.new("RGB", (20, 10), (40, 130, 90))
    img.save(path, format="HEIF")
    img.close()
    return path


@pytest.fixture
def corrupt_png_path(tmp_path: Path) -> Path:
    """Corrupt PNG file — has .png extension but contains garbage bytes, not a valid image."""
    path = tmp_path / "corrupt.png"
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 64)  # valid PNG header, but truncated/invalid body
    return path


@pytest.fixture
def fake_txt_as_png_path(tmp_path: Path) -> Path:
    """A .txt file disguised with .png extension — not an image at all."""
    path = tmp_path / "notanimage.png"
    path.write_text("This is just text, not an image.", encoding="utf-8")
    return path
