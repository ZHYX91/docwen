"""services 单元测试。"""

from __future__ import annotations

import os
import shutil
import threading
from pathlib import Path

import pytest
from PIL import Image

from docwen.services.strategies.image.merge import MergeImagesToTiffStrategy

pytestmark = pytest.mark.unit


def _write_png(path: Path) -> None:
    Image.new("RGB", (10, 10), (255, 255, 255)).save(path)


def test_merge_images_to_tiff_can_cancel(tmp_path: Path) -> None:
    a = tmp_path / "a.png"
    b = tmp_path / "b.png"
    _write_png(a)
    _write_png(b)

    ev = threading.Event()
    ev.set()

    strategy = MergeImagesToTiffStrategy()
    result = strategy.execute(
        str(a),
        options={
            "file_list": [str(a), str(b)],
            "mode": "smart",
            "cancel_event": ev,
        },
    )

    assert result.success is False


def test_merge_images_to_tiff_does_not_lock_source_files(tmp_path: Path) -> None:
    a = tmp_path / "a.png"
    b = tmp_path / "b.png"
    _write_png(a)
    _write_png(b)

    strategy = MergeImagesToTiffStrategy()
    result = strategy.execute(
        str(a),
        options={
            "file_list": [str(a), str(b)],
            "mode": "smart",
        },
    )

    assert result.success is True
    assert result.output_path is not None
    assert Path(result.output_path).exists() is True

    renamed = tmp_path / "a_renamed.png"
    a.rename(renamed)
    assert renamed.exists() is True


def test_merge_images_to_tiff_uniquifies_on_timestamp_collision(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)

    a = tmp_path / "a.png"
    b = tmp_path / "b.png"
    _write_png(a)
    _write_png(b)

    strategy = MergeImagesToTiffStrategy()
    r1 = strategy.execute(str(a), options={"file_list": [str(a), str(b)], "mode": "smart"})
    r2 = strategy.execute(str(a), options={"file_list": [str(a), str(b)], "mode": "smart"})

    assert r1.success is True
    assert r2.success is True
    assert r1.output_path is not None
    assert r2.output_path is not None
    assert r1.output_path != r2.output_path
    assert Path(r2.output_path).stem.endswith("_001") is True


def test_merge_images_to_tiff_falls_back_when_target_move_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)

    base_dir = tmp_path / "base_dir"
    base_dir.mkdir()
    bad_out = tmp_path / "bad_out"
    bad_out.mkdir()

    a = base_dir / "a.png"
    b = base_dir / "b.png"
    _write_png(a)
    _write_png(b)

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory", lambda *_args, **_kwargs: str(bad_out), raising=True
    )

    def _stub_move_file_with_retry(source: str, destination: str, *_args, **_kwargs):
        if os.path.normcase(destination).startswith(os.path.normcase(str(bad_out))):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        shutil.move(source, destination)
        return destination

    monkeypatch.setattr("docwen.utils.workspace_manager.move_file_with_retry", _stub_move_file_with_retry, raising=True)

    strategy = MergeImagesToTiffStrategy()
    result = strategy.execute(str(a), options={"file_list": [str(a), str(b)], "mode": "smart"})

    assert result.success is True
    assert result.output_path is not None
    out = Path(result.output_path)
    assert out.exists() is True
    assert os.path.normcase(str(base_dir)) in os.path.normcase(str(out.parent))
