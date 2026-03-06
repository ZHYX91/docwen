"""converter 单元测试。"""

from __future__ import annotations

import threading
from pathlib import Path

import pytest

import docwen.converter.smart_converter as smart_converter
from docwen.converter.smart_converter import OfficeSoftwareNotFoundError, SmartConverter
from docwen.utils import workspace_manager

pytestmark = pytest.mark.unit


def _touch(path: Path) -> None:
    path.write_bytes(b"x")


def test_plan_conversion_path_single_step_when_source_is_hub() -> None:
    sc = SmartConverter()
    assert sc._plan_conversion_path("xlsx", "ods", "spreadsheet") == ["ods"]


def test_plan_conversion_path_two_steps_when_neither_is_hub() -> None:
    sc = SmartConverter()
    assert sc._plan_conversion_path("xls", "ods", "spreadsheet") == ["xlsx", "ods"]


def test_plan_conversion_path_rejects_unknown_category() -> None:
    sc = SmartConverter()
    assert sc._plan_conversion_path("xls", "xlsx", "unknown") == []


def test_plan_conversion_path_rejects_unsupported_format() -> None:
    sc = SmartConverter()
    assert sc._plan_conversion_path("pdf", "xlsx", "spreadsheet") == []


def test_convert_calls_single_step(monkeypatch: pytest.MonkeyPatch) -> None:
    sc = SmartConverter()

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["ods"], raising=True)

    called = {}

    def _single(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir,
        cancel_event,
        progress_callback,
        preferred_software,
    ):
        called["source_format"] = source_format
        called["target_format"] = target_format
        return "out.ods"

    monkeypatch.setattr(SmartConverter, "_convert_single_step", _single, raising=True)

    out = sc.convert(
        input_path="in.any",
        target_format="ods",
        category="spreadsheet",
        actual_format="xlsx",
    )
    assert out == "out.ods"
    assert called == {"source_format": "xlsx", "target_format": "ods"}


def test_convert_precheck_failure_raises_office_not_found(monkeypatch: pytest.MonkeyPatch) -> None:
    """预检查失败时应抛出 OfficeSoftwareNotFoundError，而非静默返回 None"""
    sc = SmartConverter()
    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (False, "no office"), raising=True
    )
    with pytest.raises(OfficeSoftwareNotFoundError, match="no office"):
        sc.convert("in.any", "ods", "spreadsheet", "xlsx")


def test_convert_propagates_internal_error_as_runtime_error(monkeypatch: pytest.MonkeyPatch) -> None:
    """内部转换异常应包装为 RuntimeError 向上抛出，便于调用方获取失败原因"""
    sc = SmartConverter()

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["ods"], raising=True)

    def _single_raises(*_args, **_kwargs):
        raise ValueError("底层转换出错")

    monkeypatch.setattr(SmartConverter, "_convert_single_step", _single_raises, raising=True)

    with pytest.raises(RuntimeError, match="底层转换出错"):
        sc.convert("in.any", "ods", "spreadsheet", "xlsx")


def test_convert_raises_when_no_conversion_path(monkeypatch: pytest.MonkeyPatch) -> None:
    sc = SmartConverter()
    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: [], raising=True)

    with pytest.raises(RuntimeError, match="无法规划转换路径"):
        sc.convert("in.any", "ods", "spreadsheet", "xlsx")


def test_convert_single_step_stages_in_temp_then_moves_file(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    sc = SmartConverter()

    input_path = tmp_path / "in.csv"
    _touch(input_path)
    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["xlsx"], raising=True)

    seen: dict[str, str] = {}

    def _exec(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: str | None,
        cancel_event,
        exclude_wps: bool,
        preferred_software,
        original_format: str | None = None,
    ) -> str:
        assert output_dir is not None
        seen["output_dir"] = output_dir
        produced = Path(output_dir) / "out.xlsx"
        produced.write_bytes(b"x")
        return str(produced)

    monkeypatch.setattr(SmartConverter, "_execute_conversion", _exec, raising=True)

    out = sc.convert(
        input_path=str(input_path),
        target_format="xlsx",
        category="spreadsheet",
        actual_format="csv",
        output_dir=str(out_dir),
    )

    assert out is not None
    assert Path(seen["output_dir"]) != out_dir
    assert Path(out).exists() is True
    assert str(Path(out)).startswith(str(out_dir))


def test_convert_single_step_stages_in_temp_then_moves_csv_folder(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    sc = SmartConverter()

    input_path = tmp_path / "in.xlsx"
    _touch(input_path)
    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["csv"], raising=True)

    seen: dict[str, str] = {}

    def _exec(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: str | None,
        cancel_event,
        exclude_wps: bool,
        preferred_software,
        original_format: str | None = None,
    ) -> str:
        assert output_dir is not None
        seen["output_dir"] = output_dir
        csv_dir = Path(output_dir) / "csv_folder"
        csv_dir.mkdir(parents=True, exist_ok=True)
        csv_file = csv_dir / "data.csv"
        csv_file.write_text("a,b\n1,2\n", encoding="utf-8")
        return str(csv_file)

    monkeypatch.setattr(SmartConverter, "_execute_conversion", _exec, raising=True)

    out = sc.convert(
        input_path=str(input_path),
        target_format="csv",
        category="spreadsheet",
        actual_format="xlsx",
        output_dir=str(out_dir),
    )

    assert out is not None
    assert Path(seen["output_dir"]) != out_dir
    assert Path(out).exists() is True
    assert str(Path(out)).startswith(str(out_dir))
    assert Path(out).name == "data.csv"


def test_convert_single_step_respects_cancel(monkeypatch: pytest.MonkeyPatch) -> None:
    sc = SmartConverter()
    ev = threading.Event()
    ev.set()

    monkeypatch.setattr(
        SmartConverter,
        "_execute_conversion",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("should not run")),
        raising=True,
    )
    assert sc._convert_single_step("in.any", "xlsx", "ods", None, ev, None, None) is None


def test_should_exclude_wps_only_for_ods_odt() -> None:
    sc = SmartConverter()
    assert sc._should_exclude_wps("xlsx", "ods") is True
    assert sc._should_exclude_wps("odt", "docx") is True
    assert sc._should_exclude_wps("xlsx", "xls") is False
    assert sc._should_exclude_wps("docx", "rtf") is False


def test_convert_multi_step_moves_final_and_keeps_intermediate_when_enabled(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    sc = SmartConverter()

    input_path = tmp_path / "in.xls"
    _touch(input_path)

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(
        SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["xlsx", "ods"], raising=True
    )
    monkeypatch.setattr(SmartConverter, "_should_keep_intermediates", staticmethod(lambda: True), raising=True)

    def _exec(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: str | None,
        cancel_event,
        exclude_wps: bool,
        preferred_software,
        original_format: str | None = None,
    ) -> str:
        assert output_dir is not None
        name = f"{source_format}_to_{target_format}.{'ods' if target_format == 'ods' else 'xlsx'}"
        out = Path(output_dir) / name
        out.write_text(f"{source_format}->{target_format}", encoding="utf-8")
        return str(out)

    monkeypatch.setattr(SmartConverter, "_execute_conversion", _exec, raising=True)

    out = sc.convert(
        input_path=str(input_path),
        target_format="ods",
        category="spreadsheet",
        actual_format="xls",
        output_dir=str(tmp_path),
    )

    assert out is not None
    assert Path(out).exists() is True
    assert (tmp_path / "xls_to_xlsx.xlsx").exists() is True


def test_convert_multi_step_does_not_overwrite_existing_final_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    sc = SmartConverter()

    input_path = tmp_path / "in.xls"
    _touch(input_path)

    existing = tmp_path / "ods.out"
    existing.write_text("old", encoding="utf-8")

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(
        SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["xlsx", "ods"], raising=True
    )
    monkeypatch.setattr(SmartConverter, "_should_keep_intermediates", staticmethod(lambda: False), raising=True)

    def _exec(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: str | None,
        cancel_event,
        exclude_wps: bool,
        preferred_software,
        original_format: str | None = None,
    ) -> str:
        assert output_dir is not None
        out = Path(output_dir) / f"{target_format}.out"
        out.write_text(f"{source_format}->{target_format}", encoding="utf-8")
        return str(out)

    monkeypatch.setattr(SmartConverter, "_execute_conversion", _exec, raising=True)

    out = sc.convert(
        input_path=str(input_path),
        target_format="ods",
        category="spreadsheet",
        actual_format="xls",
        output_dir=str(tmp_path),
    )

    assert out is not None
    assert Path(out).name == "ods_001.out"
    assert existing.read_text(encoding="utf-8") == "old"
    assert Path(out).exists() is True


def test_convert_csv_folder_does_not_delete_existing_target_folder(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    sc = SmartConverter()

    input_path = tmp_path / "in.xlsx"
    _touch(input_path)

    existing_folder = tmp_path / "csv_folder"
    existing_folder.mkdir()
    (existing_folder / "keep.txt").write_text("old", encoding="utf-8")

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(
        SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["xlsx", "csv"], raising=True
    )
    monkeypatch.setattr(SmartConverter, "_should_keep_intermediates", staticmethod(lambda: False), raising=True)

    def _exec(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: str | None,
        cancel_event,
        exclude_wps: bool,
        preferred_software,
        original_format: str | None = None,
    ) -> str:
        assert output_dir is not None
        if target_format == "xlsx":
            out = Path(output_dir) / "xlsx.out"
            out.write_bytes(b"x")
            return str(out)

        csv_dir = Path(output_dir) / "csv_folder"
        csv_dir.mkdir(parents=True, exist_ok=True)
        csv_file = csv_dir / "data.csv"
        csv_file.write_text("a,b\n1,2\n", encoding="utf-8")
        return str(csv_file)

    monkeypatch.setattr(SmartConverter, "_execute_conversion", _exec, raising=True)

    out = sc.convert(
        input_path=str(input_path),
        target_format="csv",
        category="spreadsheet",
        actual_format="xlsx",
        output_dir=str(tmp_path),
    )

    assert out is not None
    assert existing_folder.exists() is True
    assert (existing_folder / "keep.txt").read_text(encoding="utf-8") == "old"
    assert Path(out).name == "data.csv"
    assert "csv_folder_001" in str(out)
    assert Path(out).exists() is True


def test_convert_multi_step_progress_callback_called_for_each_step(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    sc = SmartConverter()

    input_path = tmp_path / "in.xls"
    _touch(input_path)

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(
        SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["xlsx", "ods"], raising=True
    )
    monkeypatch.setattr(SmartConverter, "_should_keep_intermediates", staticmethod(lambda: False), raising=True)
    monkeypatch.setattr(smart_converter, "t", lambda *_args, **_kwargs: "msg", raising=True)

    calls = []

    def _progress(msg: str) -> None:
        calls.append(msg)

    def _exec(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: str | None,
        cancel_event,
        exclude_wps: bool,
        preferred_software,
        original_format: str | None = None,
    ) -> str:
        assert output_dir is not None
        out = Path(output_dir) / f"{target_format}.out"
        out.write_bytes(b"x")
        return str(out)

    monkeypatch.setattr(SmartConverter, "_execute_conversion", _exec, raising=True)

    out = sc.convert(
        input_path=str(input_path),
        target_format="ods",
        category="spreadsheet",
        actual_format="xls",
        output_dir=str(tmp_path),
        progress_callback=_progress,
    )

    assert out is not None
    assert len(calls) == 2


def test_convert_multi_step_uses_fallback_when_output_dir_fails(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    sc = SmartConverter()

    input_path = tmp_path / "in.xls"
    _touch(input_path)

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True
    )
    monkeypatch.setattr(
        SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["xlsx", "ods"], raising=True
    )
    monkeypatch.setattr(SmartConverter, "_should_keep_intermediates", staticmethod(lambda: False), raising=True)

    def _exec(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: str | None,
        cancel_event,
        exclude_wps: bool,
        preferred_software,
        original_format: str | None = None,
    ) -> str:
        assert output_dir is not None
        out = Path(output_dir) / f"{target_format}.out"
        out.write_text(f"{source_format}->{target_format}", encoding="utf-8")
        return str(out)

    monkeypatch.setattr(SmartConverter, "_execute_conversion", _exec, raising=True)

    def _fake_move_with_retry(source: str, destination: str, max_retries: int = 3, retry_delay: float = 0.5) -> str | None:
        if str(Path(destination)).startswith(str(out_dir)):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        import shutil

        shutil.move(source, destination)
        return destination

    monkeypatch.setattr(workspace_manager, "move_file_with_retry", _fake_move_with_retry, raising=True)

    out = sc.convert(
        input_path=str(input_path),
        target_format="ods",
        category="spreadsheet",
        actual_format="xls",
        output_dir=str(out_dir),
    )

    assert out is not None
    assert Path(out).exists() is True
    assert str(Path(out)).startswith(str(tmp_path))
    assert str(Path(out)).startswith(str(out_dir)) is False
