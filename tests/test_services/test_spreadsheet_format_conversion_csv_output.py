"""services 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.services.strategies import get_strategy

pytestmark = pytest.mark.unit


def test_format_conversion_xls_to_csv_moves_file_when_converter_returns_file_in_temp_root(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    src = tmp_path / "a.xls"
    src.write_bytes(b"x")

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )

    def _stub_convert(
        _self,
        *,
        input_path: str,
        target_format: str,
        category: str,
        actual_format: str,
        output_dir: str,
        cancel_event=None,
        progress_callback=None,
        preferred_software=None,
    ) -> str:
        p = Path(output_dir) / "out.csv"
        p.write_text("1,2\n", encoding="utf-8")
        return str(p)

    monkeypatch.setattr("docwen.converter.smart_converter.SmartConverter.convert", _stub_convert, raising=True)

    strategy_cls = get_strategy(action_type=None, source_format="xls", target_format="csv")
    strategy = strategy_cls()

    result = strategy.execute(str(src), options={"actual_format": "xls"})

    assert result.success is True
    assert result.output_path is not None
    output_path = Path(result.output_path)
    assert output_path.exists() is True
    assert output_path.parent == out_dir
    assert output_path.name == "out.csv"
