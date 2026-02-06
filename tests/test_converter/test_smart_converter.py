from __future__ import annotations

import threading
from pathlib import Path

import pytest

from docwen.converter.smart_converter import SmartConverter
import docwen.converter.smart_converter as smart_converter


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

    monkeypatch.setattr(smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True)
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


def test_convert_precheck_failure_returns_none(monkeypatch: pytest.MonkeyPatch) -> None:
    sc = SmartConverter()
    monkeypatch.setattr(smart_converter, "check_office_availability", lambda *_args, **_kwargs: (False, "no office"), raising=True)
    assert sc.convert("in.any", "ods", "spreadsheet", "xlsx") is None


def test_convert_single_step_respects_cancel(monkeypatch: pytest.MonkeyPatch) -> None:
    sc = SmartConverter()
    ev = threading.Event()
    ev.set()

    monkeypatch.setattr(SmartConverter, "_execute_conversion", lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("should not run")), raising=True)
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

    monkeypatch.setattr(smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True)
    monkeypatch.setattr(SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["xlsx", "ods"], raising=True)
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


def test_convert_multi_step_progress_callback_called_for_each_step(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    sc = SmartConverter()

    input_path = tmp_path / "in.xls"
    _touch(input_path)

    monkeypatch.setattr(smart_converter, "check_office_availability", lambda *_args, **_kwargs: (True, ""), raising=True)
    monkeypatch.setattr(SmartConverter, "_plan_conversion_path", lambda *_args, **_kwargs: ["xlsx", "ods"], raising=True)
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

