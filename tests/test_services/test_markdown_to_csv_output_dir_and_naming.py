"""services 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.services.strategies import get_strategy

pytestmark = pytest.mark.unit


def test_markdown_to_csv_ignores_source_dir_collision_when_output_dir_is_custom(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    src = src_dir / "a_20200101_010101_fromMd.md"
    src.write_text("# hi\n", encoding="utf-8")

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.path_utils.generate_timestamp",
        lambda: "20200101_010101",
        raising=True,
    )

    (src_dir / "a_20200101_010101_fromMd").mkdir()

    def _stub_convert_md_to_xlsx(
        md_path: str,
        output_path: str,
        *,
        template_name: str,
        progress_callback=None,
        cancel_event=None,
        original_source_path: str | None = None,
    ) -> str:
        Path(output_path).write_bytes(b"x")
        return output_path

    monkeypatch.setattr(
        "docwen.services.strategies.markdown.to_spreadsheet.convert_md_to_xlsx",
        _stub_convert_md_to_xlsx,
        raising=True,
    )

    def _stub_xlsx_to_csv(
        xlsx_path: str,
        *,
        actual_format=None,
        output_dir: str,
        original_basename: str,
        unified_timestamp_desc: str,
    ) -> list[str]:
        assert original_basename == "a"
        assert unified_timestamp_desc == "20200101_010101_fromMd"
        p = Path(output_dir) / f"{original_basename}_{unified_timestamp_desc}.csv"
        p.write_text("1,2\n", encoding="utf-8")
        return [str(p)]

    monkeypatch.setattr(
        "docwen.converter.formats.spreadsheet.xlsx_to_csv",
        _stub_xlsx_to_csv,
        raising=True,
    )

    strategy_cls = get_strategy(action_type=None, source_format="md", target_format="csv")
    strategy = strategy_cls()

    result = strategy.execute(str(src), options={"template_name": "t"})

    assert result.success is True
    assert result.output_path is not None

    output_path = Path(result.output_path)
    assert output_path.exists() is True
    assert output_path.parent.parent == out_dir
    assert output_path.parent.name == "a_20200101_010101_fromMd"
