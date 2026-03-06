"""services 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.services.strategies.layout.to_image import LayoutToPngStrategy, LayoutToTifStrategy
from docwen.utils import workspace_manager

pytestmark = pytest.mark.unit


def _create_pdf(path: Path, pages: int) -> None:
    fitz = pytest.importorskip("fitz")
    doc = fitz.open()
    for _ in range(pages):
        doc.new_page()
    doc.save(str(path))
    doc.close()


def test_layout_to_png_renders_one_image_per_page(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pytest.importorskip("fitz")

    src = tmp_path / "src.pdf"
    _create_pdf(src, pages=2)

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = LayoutToPngStrategy().execute(
        str(src),
        options={"actual_format": "pdf", "dpi": 72},
    )

    assert result.success is True
    assert result.output_path is not None

    out_folder = Path(result.output_path)
    assert out_folder.exists()
    assert len(list(out_folder.glob("*.png"))) == 2


def test_layout_to_tif_renders_multi_page_tif(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pytest.importorskip("fitz")
    try:
        from PIL import Image
    except ImportError:
        pytest.skip("PIL not installed")

    src = tmp_path / "src.pdf"
    _create_pdf(src, pages=2)

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = LayoutToTifStrategy().execute(
        str(src),
        options={"actual_format": "pdf", "dpi": 72},
    )

    assert result.success is True
    assert result.output_path is not None

    out_tif = Path(result.output_path)
    assert out_tif.exists()

    with Image.open(out_tif) as img:
        assert getattr(img, "n_frames", 1) == 2
