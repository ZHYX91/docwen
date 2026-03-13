"""utils 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.utils import workspace_manager


pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_get_output_directory_prefers_custom_dir_param(tmp_path: Path) -> None:
    input_file = tmp_path / "in.docx"
    input_file.write_text("x", encoding="utf-8")

    custom_dir = tmp_path / "custom"
    out = workspace_manager.get_output_directory(str(input_file), custom_dir=str(custom_dir))
    assert out == str(custom_dir)


@pytest.mark.unit
def test_get_output_directory_uses_source_dir_when_config_unavailable(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_file = tmp_path / "in.docx"
    input_file.write_text("x", encoding="utf-8")

    from docwen.config.config_manager import config_manager

    monkeypatch.setattr(
        config_manager, "get_output_directory_settings", lambda: (_ for _ in ()).throw(RuntimeError("boom"))
    )

    out = workspace_manager.get_output_directory(str(input_file))
    assert out == str(tmp_path)


@pytest.mark.unit
def test_get_output_directory_custom_mode_creates_date_subfolder(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_file = tmp_path / "in.docx"
    input_file.write_text("x", encoding="utf-8")

    out_base = tmp_path / "out"
    settings = {
        "mode": "custom",
        "custom_path": str(out_base),
        "create_date_subfolder": True,
        "date_folder_format": "%Y-%m-%d",
    }

    from docwen.config.config_manager import config_manager

    monkeypatch.setattr(config_manager, "get_output_directory_settings", lambda: settings)

    class _FixedDatetime:
        @staticmethod
        def now():
            import datetime as _dt

            return _dt.datetime(2020, 1, 2, 3, 4, 5)

    monkeypatch.setattr(workspace_manager, "datetime", _FixedDatetime)

    out = workspace_manager.get_output_directory(str(input_file))
    assert out == str(out_base / "2020-01-02")
    assert Path(out).exists()
    assert Path(out).is_dir()
