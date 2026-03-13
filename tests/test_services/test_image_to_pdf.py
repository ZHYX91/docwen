"""services 单元测试。"""

from __future__ import annotations

import contextlib
from pathlib import Path

import pytest
from PIL import Image

from docwen.services.strategies.image.to_pdf import ImageToPdfStrategy
from docwen.services.strategies.image.utils import extract_tiff_pages, is_multipage_tiff
from docwen.utils import workspace_manager

pytestmark = pytest.mark.unit


def test_image_to_pdf_multipage_tiff(tmp_path: Path) -> None:
    try:
        import fitz  # type: ignore
    except Exception:
        pytest.skip("PyMuPDF not available")

    src = tmp_path / "in.tif"
    frames = [
        Image.new("RGB", (64, 64), (255, 255, 255)),
        Image.new("RGB", (64, 64), (128, 128, 128)),
        Image.new("RGB", (64, 64), (0, 0, 0)),
    ]
    frames[0].save(src, save_all=True, append_images=frames[1:], format="TIFF", compression="tiff_lzw")
    for f in frames:
        with contextlib.suppress(Exception):
            f.close()

    strategy = ImageToPdfStrategy()
    result = strategy.execute(
        str(src),
        options={"quality_mode": "original", "actual_format": "tif"},
    )

    assert result.success is True
    assert result.output_path is not None
    pdf_path = Path(result.output_path)
    assert pdf_path.exists() is True

    with fitz.open(pdf_path) as doc:
        assert doc.page_count == 3


def test_image_to_pdf_multipage_tiff_without_actual_format(tmp_path: Path) -> None:
    try:
        import fitz  # type: ignore
    except Exception:
        pytest.skip("PyMuPDF not available")

    src = tmp_path / "in.tif"
    frames = [
        Image.new("RGB", (64, 64), (255, 255, 255)),
        Image.new("RGB", (64, 64), (128, 128, 128)),
        Image.new("RGB", (64, 64), (0, 0, 0)),
    ]
    frames[0].save(src, save_all=True, append_images=frames[1:], format="TIFF", compression="tiff_lzw")
    for f in frames:
        with contextlib.suppress(Exception):
            f.close()

    strategy = ImageToPdfStrategy()
    result = strategy.execute(
        str(src),
        options={"quality_mode": "original"},
    )

    assert result.success is True
    assert result.output_path is not None
    pdf_path = Path(result.output_path)
    assert pdf_path.exists() is True

    with fitz.open(pdf_path) as doc:
        assert doc.page_count == 3


def test_image_to_pdf_bmp_is_preprocessed_to_png(tmp_path: Path) -> None:
    src = tmp_path / "in.bmp"
    img = Image.new("RGB", (64, 64), (255, 255, 255))
    img.save(src, format="BMP")
    img.close()

    strategy = ImageToPdfStrategy()
    result = strategy.execute(
        str(src),
        options={"quality_mode": "original", "actual_format": "bmp"},
    )

    assert result.success is True
    assert result.output_path is not None
    pdf_path = Path(result.output_path)
    assert pdf_path.exists() is True
    assert pdf_path.suffix.lower() == ".pdf"


def test_is_multipage_tiff_works_with_wrong_extension(tmp_path: Path) -> None:
    src = tmp_path / "in.tif"
    frames = [
        Image.new("RGB", (32, 32), (255, 255, 255)),
        Image.new("RGB", (32, 32), (0, 0, 0)),
    ]
    frames[0].save(src, save_all=True, append_images=frames[1:], format="TIFF", compression="tiff_lzw")
    for f in frames:
        with contextlib.suppress(Exception):
            f.close()

    renamed = tmp_path / "in.jpg"
    src.rename(renamed)

    assert is_multipage_tiff(str(renamed)) is True


def test_extract_tiff_pages_composites_transparency_on_white(tmp_path: Path) -> None:
    src = tmp_path / "in.tif"

    rgba = Image.new("RGBA", (10, 10), (255, 0, 0, 0))
    rgba.putpixel((1, 1), (255, 0, 0, 255))

    la = Image.new("LA", (10, 10), (0, 0))
    la.putpixel((1, 1), (0, 255))

    rgba.save(src, save_all=True, append_images=[la], format="TIFF", compression="tiff_lzw")
    for f in (rgba, la):
        with contextlib.suppress(Exception):
            f.close()

    pages = extract_tiff_pages(str(src), output_dir=str(tmp_path), actual_format="tiff")
    assert len(pages) == 2

    p1 = Image.open(pages[0][1])
    try:
        assert p1.mode == "RGB"
        assert p1.getpixel((0, 0)) == (255, 255, 255)
        assert p1.getpixel((1, 1)) == (255, 0, 0)
    finally:
        p1.close()

    p2 = Image.open(pages[1][1])
    try:
        assert p2.mode == "RGB"
        assert p2.getpixel((0, 0)) == (255, 255, 255)
        assert p2.getpixel((1, 1)) == (0, 0, 0)
    finally:
        p2.close()


def test_image_to_pdf_keep_intermediates_still_saves_final_pdf_when_output_dir_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    src = tmp_path / "in.tif"
    frames = [
        Image.new("RGB", (64, 64), (255, 255, 255)),
        Image.new("RGB", (64, 64), (0, 0, 0)),
    ]
    frames[0].save(src, save_all=True, append_images=frames[1:], format="TIFF", compression="tiff_lzw")
    for f in frames:
        with contextlib.suppress(Exception):
            f.close()

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda *_args, **_kwargs: str(out_dir), raising=True)

    monkeypatch.setattr(ImageToPdfStrategy, "_should_keep_intermediates", lambda *_args, **_kwargs: True, raising=True)

    def _fake_move_with_retry(
        source: str, destination: str, max_retries: int = 3, retry_delay: float = 0.5
    ) -> str | None:
        if str(Path(destination)).startswith(str(out_dir)):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        import shutil

        shutil.move(source, destination)
        return destination

    monkeypatch.setattr(workspace_manager, "move_file_with_retry", _fake_move_with_retry, raising=True)

    strategy = ImageToPdfStrategy()
    result = strategy.execute(
        str(src),
        options={"quality_mode": "original", "actual_format": "tif"},
    )

    assert result.success is True
    assert result.output_path is not None
    pdf_path = Path(result.output_path)
    assert pdf_path.exists() is True
    assert str(pdf_path).startswith(str(out_dir)) is False
