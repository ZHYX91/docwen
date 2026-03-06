"""services 单元测试。"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

import pytest
from PIL import Image

from docwen.services.strategies.image.format_conversion import ImageFormatConversionStrategy

pytestmark = pytest.mark.unit


def _write_png(path: Path) -> None:
    Image.new("RGB", (10, 10), (255, 255, 255)).save(path)


def test_image_format_conversion_uniquifies_on_timestamp_collision(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)

    src = tmp_path / "a.png"
    _write_png(src)

    strategy = ImageFormatConversionStrategy()
    r1 = strategy.execute(str(src), options={"actual_format": "png", "target_format": "jpg"})
    r2 = strategy.execute(str(src), options={"actual_format": "png", "target_format": "jpg"})

    assert r1.success is True
    assert r2.success is True
    assert r1.output_path is not None
    assert r2.output_path is not None
    assert Path(r1.output_path).exists() is True
    assert Path(r2.output_path).exists() is True
    assert r1.output_path != r2.output_path
    assert Path(r2.output_path).stem.endswith("_001") is True


def test_image_format_conversion_falls_back_when_target_move_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)

    base_dir = tmp_path / "base_dir"
    base_dir.mkdir()
    bad_out = tmp_path / "bad_out"
    bad_out.mkdir()

    src = base_dir / "a.png"
    _write_png(src)

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory", lambda *_a, **_k: str(bad_out), raising=True
    )

    def _stub_move_file_with_retry(source: str, destination: str, *_args, **_kwargs):
        if os.path.normcase(destination).startswith(os.path.normcase(str(bad_out))):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        shutil.move(source, destination)
        return destination

    monkeypatch.setattr("docwen.utils.workspace_manager.move_file_with_retry", _stub_move_file_with_retry, raising=True)

    strategy = ImageFormatConversionStrategy()
    result = strategy.execute(str(src), options={"actual_format": "png", "target_format": "jpg"})

    assert result.success is True
    assert result.output_path is not None
    out = Path(result.output_path)
    assert out.exists() is True
    assert os.path.normcase(str(base_dir)) in os.path.normcase(str(out.parent))
