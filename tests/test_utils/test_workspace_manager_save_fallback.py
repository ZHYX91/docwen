"""utils 单元测试。"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from docwen.utils import workspace_manager


pytestmark = pytest.mark.unit

@pytest.mark.unit
def test_save_output_with_fallback_uses_source_dir_when_target_fails(tmp_path, monkeypatch) -> None:
    temp_file = tmp_path / "temp.txt"
    temp_file.write_text("hello", encoding="utf-8")

    original_input_file = tmp_path / "input.md"
    original_input_file.write_text("# input", encoding="utf-8")

    target_path = tmp_path / "out" / "result.txt"
    expected_fallback_path = tmp_path / "result.txt"

    def fake_move_file_with_retry(source: str, destination: str, max_retries: int = 3, retry_delay: float = 0.5):
        if destination == str(target_path):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        shutil.move(source, destination)
        return destination

    monkeypatch.setattr(workspace_manager, "move_file_with_retry", fake_move_file_with_retry)

    saved_path, location = workspace_manager.save_output_with_fallback(
        str(temp_file),
        str(target_path),
        original_input_file=str(original_input_file),
    )

    assert saved_path == str(expected_fallback_path)
    assert "原文件所在目录" in location
    assert not temp_file.exists()
    assert expected_fallback_path.exists()


@pytest.mark.unit
def test_save_output_with_fallback_directory_rescue_is_unique_and_does_not_merge(tmp_path, monkeypatch) -> None:
    temp_dir = tmp_path / "tempdir"
    temp_dir.mkdir()
    (temp_dir / "a.txt").write_text("new", encoding="utf-8")

    original_input_file = tmp_path / "input.md"
    original_input_file.write_text("# input", encoding="utf-8")

    target_path = tmp_path / "out" / "result_folder"

    monkeypatch.setattr(workspace_manager, "move_file_with_retry", lambda *_args, **_kwargs: None)
    monkeypatch.setattr("tempfile.gettempdir", lambda: str(tmp_path / "sys_temp"), raising=True)

    home = tmp_path / "home"
    desktop_path = home / "Desktop"
    monkeypatch.setattr(workspace_manager.Path, "home", staticmethod(lambda: home))

    real_exists = workspace_manager.Path.exists

    def fake_exists(self) -> bool:
        if self == desktop_path:
            return False
        return real_exists(self)

    monkeypatch.setattr(workspace_manager.Path, "exists", fake_exists)

    rescue_base = tmp_path / "sys_temp" / "docwen_rescue"
    existing = rescue_base / "result_folder"
    existing.mkdir(parents=True, exist_ok=True)
    (existing / "keep.txt").write_text("old", encoding="utf-8")

    saved_path, location = workspace_manager.save_output_with_fallback(
        str(temp_dir),
        str(target_path),
        original_input_file=str(original_input_file),
    )

    assert saved_path is not None
    assert "临时救援目录" in location
    assert temp_dir.exists() is False

    saved = Path(saved_path)
    assert saved != existing
    assert (existing / "keep.txt").read_text(encoding="utf-8") == "old"
    assert (saved / "a.txt").read_text(encoding="utf-8") == "new"


@pytest.mark.unit
def test_save_output_with_fallback_file_rescue_is_unique_and_removes_source(tmp_path, monkeypatch) -> None:
    temp_file = tmp_path / "temp.txt"
    temp_file.write_text("new", encoding="utf-8")

    original_input_file = tmp_path / "input.md"
    original_input_file.write_text("# input", encoding="utf-8")

    target_path = tmp_path / "out" / "result.txt"

    monkeypatch.setattr(workspace_manager, "move_file_with_retry", lambda *_args, **_kwargs: None)
    monkeypatch.setattr("tempfile.gettempdir", lambda: str(tmp_path / "sys_temp"), raising=True)

    home = tmp_path / "home"
    desktop_path = home / "Desktop"
    monkeypatch.setattr(workspace_manager.Path, "home", staticmethod(lambda: home))

    real_exists = workspace_manager.Path.exists

    def fake_exists(self) -> bool:
        if self == desktop_path:
            return False
        return real_exists(self)

    monkeypatch.setattr(workspace_manager.Path, "exists", fake_exists)

    rescue_base = tmp_path / "sys_temp" / "docwen_rescue"
    rescue_base.mkdir(parents=True, exist_ok=True)
    existing = rescue_base / "result.txt"
    existing.write_text("old", encoding="utf-8")

    saved_path, location = workspace_manager.save_output_with_fallback(
        str(temp_file),
        str(target_path),
        original_input_file=str(original_input_file),
    )

    assert saved_path is not None
    assert "临时救援目录" in location
    assert temp_file.exists() is False
    assert existing.read_text(encoding="utf-8") == "old"
    assert Path(saved_path).read_text(encoding="utf-8") == "new"


@pytest.mark.unit
def test_save_intermediate_item_does_not_overwrite_existing_file(tmp_path: Path) -> None:
    out_dir = tmp_path / "out"
    out_dir.mkdir()

    existing = out_dir / "temp_output.xlsx"
    existing.write_text("old", encoding="utf-8")

    src = tmp_path / "temp_output.xlsx"
    src.write_text("new", encoding="utf-8")

    saved = workspace_manager.save_intermediate_item(str(src), str(out_dir), move=False)

    assert saved is not None
    assert Path(saved).name == "temp_output_001.xlsx"
    assert existing.read_text(encoding="utf-8") == "old"
    assert Path(saved).read_text(encoding="utf-8") == "new"


@pytest.mark.unit
def test_save_output_with_fallback_raises_when_rescue_copy_fails(tmp_path: Path, monkeypatch) -> None:
    temp_file = tmp_path / "temp.txt"
    temp_file.write_text("new", encoding="utf-8")

    original_input_file = tmp_path / "input.md"
    original_input_file.write_text("# input", encoding="utf-8")

    target_path = tmp_path / "out" / "result.txt"

    monkeypatch.setattr(workspace_manager, "move_file_with_retry", lambda *_args, **_kwargs: None)
    monkeypatch.setattr("tempfile.gettempdir", lambda: str(tmp_path / "sys_temp"), raising=True)

    home = tmp_path / "home"
    desktop_path = home / "Desktop"
    monkeypatch.setattr(workspace_manager.Path, "home", staticmethod(lambda: home))

    real_exists = workspace_manager.Path.exists

    def fake_exists(self) -> bool:
        if self == desktop_path:
            return False
        return real_exists(self)

    monkeypatch.setattr(workspace_manager.Path, "exists", fake_exists)

    monkeypatch.setattr(shutil, "copy2", lambda *_args, **_kwargs: (_ for _ in ()).throw(PermissionError("nope")))

    with pytest.raises(RuntimeError, match="救援目录"):
        workspace_manager.save_output_with_fallback(
            str(temp_file),
            str(target_path),
            original_input_file=str(original_input_file),
        )
