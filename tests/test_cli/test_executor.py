"""CLI 单元测试。"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from PIL import Image

import docwen.cli.executor as executor
from docwen.errors import InvalidInputError
from docwen.services.error_codes import ERROR_CODE_INVALID_INPUT
from docwen.services.result import ConversionResult

pytestmark = pytest.mark.unit


class _OkStrategy:
    def execute(self, *, file_path: str, options: dict, progress_callback=None):
        return ConversionResult(success=True, output_path="out", message="ok")


class _FailStrategy:
    def execute(self, *, file_path: str, options: dict, progress_callback=None):
        raise RuntimeError("boom")


def test_get_strategy_for_action_convert_to_md_uses_category(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    called = {}

    def _stub_get_strategy(*, action_type=None, source_format=None, target_format=None):
        called["action_type"] = action_type
        called["source_format"] = source_format
        called["target_format"] = target_format
        return _OkStrategy

    monkeypatch.setattr(executor, "get_strategy", _stub_get_strategy, raising=True)
    monkeypatch.setattr(
        "docwen.utils.file_type_utils.ACTUAL_FORMAT_TO_CATEGORY",
        {"docx": "document"},
        raising=False,
    )

    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    assert (
        executor.get_strategy_for_action("convert", str(f), {"actual_format": "docx", "target_format": "md"})
        is _OkStrategy
    )
    assert called == {"action_type": None, "source_format": "document", "target_format": "md"}


def test_get_supported_convert_targets_includes_md_from_category_registry(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("docwen.services.strategies.load_all", lambda: None, raising=True)
    monkeypatch.setattr(
        "docwen.services.strategies.get_conversion_registry",
        lambda: {("document", "md"): object, ("docx", "pdf"): object},
        raising=True,
    )

    targets = executor.get_supported_convert_targets()
    assert "md" in targets


def test_get_strategy_for_action_convert_requires_target(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(executor, "get_strategy", lambda **_kwargs: _OkStrategy, raising=True)

    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    with pytest.raises(InvalidInputError) as excinfo:
        executor.get_strategy_for_action("convert", str(f), {"actual_format": "docx"})

    assert "需要指定目标格式" in str(excinfo.value)


def test_execute_action_missing_file_returns_1_text(capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    missing = tmp_path / "missing.docx"
    exit_code = executor.execute_action("convert", str(missing), json_mode=False)
    captured = capsys.readouterr()

    assert exit_code == 2
    assert "文件不存在" in captured.err


def test_execute_action_missing_file_returns_1_json(capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    missing = tmp_path / "missing.docx"
    exit_code = executor.execute_action("convert", str(missing), json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 2
    payload = json.loads(out)
    assert payload["success"] is False
    assert payload["command"] == "convert"
    assert "文件不存在" in (payload.get("error") or {}).get("message", "")
    assert (payload.get("error") or {}).get("error_code") == ERROR_CODE_INVALID_INPUT


def test_execute_action_actual_format_detection_failure_is_non_fatal(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    monkeypatch.setattr(
        "docwen.utils.file_type_utils.detect_actual_file_format",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("x")),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.services.strategies.get_strategy",
        lambda **_kwargs: _OkStrategy,
        raising=True,
    )

    assert executor.execute_action("convert", str(f), options={"target_format": "md"}, json_mode=True) == 0


def test_execute_action_convert_to_md_image_detection_failure_still_succeeds(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    img = tmp_path / "a.png"
    Image.new("RGB", (10, 10), (255, 255, 255)).save(img)

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.file_type_utils.detect_actual_file_format",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("x")),
        raising=True,
    )

    class _Strategy:
        def execute(self, *, file_path: str, options: dict, progress_callback=None):
            out_file = out_dir / "out.md"
            out_file.write_text("ok", encoding="utf-8")
            return ConversionResult(success=True, output_path=str(out_file), message="ok")

    monkeypatch.setattr("docwen.services.strategies.get_strategy", lambda **_kwargs: _Strategy, raising=True)

    exit_code = executor.execute_action(
        "convert",
        str(img),
        options={"target_format": "md", "extract_ocr": False},
        json_mode=True,
    )
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["success"] is True
    assert Path(payload["data"]["output_file"]).exists()


def test_execute_action_strategy_exception_returns_1_json(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    monkeypatch.setattr(
        "docwen.services.strategies.get_strategy",
        lambda **_kwargs: _FailStrategy,
        raising=True,
    )

    exit_code = executor.execute_action("convert", str(f), options={"target_format": "md"}, json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 1
    payload = json.loads(out)
    assert payload["success"] is False
    assert payload["command"] == "convert"
    assert "boom" in ((payload.get("error") or {}).get("message") or "")


def test_get_supported_actions_covers_categories() -> None:
    markdown = executor.get_supported_actions("markdown", "md")
    assert {a["name"] for a in markdown} == {"convert"}

    document = executor.get_supported_actions("document", "docx")
    assert {a["name"] for a in document} == {"convert", "validate"}

    spreadsheet = executor.get_supported_actions("spreadsheet", "xlsx")
    assert {a["name"] for a in spreadsheet} == {"convert", "merge_tables"}

    layout = executor.get_supported_actions("layout", "pdf")
    assert {a["name"] for a in layout} == {"convert", "merge_pdfs", "split_pdf"}

    image = executor.get_supported_actions("image", "png")
    assert {a["name"] for a in image} == {"convert"}

    assert executor.get_supported_actions("unknown", "bin") == []


def test_inspect_file_json(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    monkeypatch.setattr("docwen.cli.utils.detect_category", lambda *_args, **_kwargs: "markdown", raising=True)
    monkeypatch.setattr("docwen.cli.utils.detect_format", lambda *_args, **_kwargs: "md", raising=True)

    exit_code = executor.inspect_file("x.md", json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["data"]["category"] == "markdown"
    assert payload["data"]["format"] == "md"
    assert {a["name"] for a in payload["data"]["supported_actions"]} == {"convert"}


def test_list_all_actions_json(capsys: pytest.CaptureFixture[str]) -> None:
    exit_code = executor.list_all_actions(json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["command"] == "actions"
    assert "actions" in payload["data"]
    assert any(a["name"] == "convert" for a in payload["data"]["actions"])


def test_list_numbering_schemes_fallback_json(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    import importlib

    cm = importlib.import_module("docwen.config.config_manager")
    monkeypatch.setattr(
        cm.config_manager, "get_heading_schemes", lambda: (_ for _ in ()).throw(RuntimeError("x")), raising=True
    )

    exit_code = executor.list_numbering_schemes(json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["command"] == "numbering-schemes"
    assert "schemes" in payload["data"]
    assert any(s["id"] == "gongwen_standard" for s in payload["data"]["schemes"])


def test_list_templates_error_json(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    import docwen.template.loader as loader

    class _BadLoader:
        def __init__(self):
            raise RuntimeError("no templates")

    monkeypatch.setattr(loader, "TemplateLoader", _BadLoader, raising=True)

    exit_code = executor.list_templates(json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 1
    payload = json.loads(out)
    assert payload["success"] is False
    assert payload["command"] == "templates"
    assert "no templates" in ((payload.get("error") or {}).get("message") or "")


def test_list_templates_json(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    import docwen.template.loader as loader

    class _OkLoader:
        def get_available_templates(self, target: str):
            if target == "docx":
                return ["a", "b"]
            if target == "xlsx":
                return ["c"]
            return []

    monkeypatch.setattr(loader, "TemplateLoader", _OkLoader, raising=True)

    exit_code = executor.list_templates(json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["command"] == "templates"
    assert isinstance(payload["data"].get("templates"), list)
    assert {"id": "a", "name": "a", "target": "docx", "description": None, "example": None} in payload["data"][
        "templates"
    ]
    assert {"id": "c", "name": "c", "target": "xlsx", "description": None, "example": None} in payload["data"][
        "templates"
    ]


def test_list_optimizations_json(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    import importlib

    cm = importlib.import_module("docwen.config.config_manager")
    monkeypatch.setattr(
        cm.config_manager,
        "get_optimization_types",
        lambda: {
            "x": {"enabled": True, "name": "X", "description": "D", "scopes": ["document_to_md"], "locales": ["*"]}
        },
        raising=True,
    )
    monkeypatch.setattr(cm.config_manager, "get_optimization_settings", lambda: {"order": ["x"]}, raising=True)

    exit_code = executor.list_optimizations(json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["command"] == "optimizations"
    assert payload["data"]["optimizations"][0]["id"] == "x"
    assert payload["data"]["optimizations"][0]["examples"] == []
    assert payload["data"]["optimizations"][0]["recommended_for"] == []


def test_list_formats_json(capsys: pytest.CaptureFixture[str]) -> None:
    exit_code = executor.list_formats(json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["command"] == "formats"
    assert isinstance(payload["data"].get("formats"), list)
    assert any("targets" in item and "source" in item for item in payload["data"]["formats"])


def test_execute_batch_interrupt_returns_130_and_sets_json_fields(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    def _stub_execute_batch(self, *_args, **_kwargs):
        raise KeyboardInterrupt()

    monkeypatch.setattr(executor.ConversionService, "execute_batch", _stub_execute_batch, raising=True)

    exit_code = executor.execute_batch(
        "convert",
        ["a.docx", "b.docx"],
        options={},
        json_mode=True,
        continue_on_error=True,
    )
    out = capsys.readouterr().out

    assert exit_code == 130
    payload = json.loads(out)
    assert payload["command"] == "convert"
    assert payload["data"]["interrupted"] is True
    assert payload["data"]["total"] == 2
    assert payload["data"]["processed_count"] == 1
    assert payload["data"]["results"][0]["error"] == "用户中断"


def test_execute_batch_failure_breaks_and_reports_processed_count(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen.services.batch import BatchResult, FileResult

    def _stub_execute_batch(self, batch_request, **_kwargs):
        return BatchResult(
            total=len(batch_request.requests),
            processed_count=1,
            results=[
                FileResult(
                    file_path=batch_request.requests[0].file_path or "",
                    success=False,
                    status="failed",
                    message="nope",
                    error_code=None,
                    details=None,
                )
            ],
            success_count=0,
            failed_count=1,
            cancelled=False,
        )

    monkeypatch.setattr(executor.ConversionService, "execute_batch", _stub_execute_batch, raising=True)

    exit_code = executor.execute_batch(
        "convert",
        ["a.docx", "b.docx", "c.docx"],
        options={},
        json_mode=True,
        continue_on_error=False,
    )
    out = capsys.readouterr().out

    assert exit_code == 1
    payload = json.loads(out)
    assert payload["command"] == "convert"
    assert payload["data"]["interrupted"] is False
    assert payload["data"]["total"] == 3
    assert payload["data"]["processed_count"] == 1
    assert payload["data"]["failed_count"] == 1


def test_execute_batch_json_is_single_document_and_options_are_isolated(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen.services.batch import BatchResult, FileResult

    def _stub_execute_batch(self, batch_request, **_kwargs):
        for req in batch_request.requests:
            assert req.options.get("marker") is None
            req.options["marker"] = req.file_path
        return BatchResult(
            total=len(batch_request.requests),
            processed_count=len(batch_request.requests),
            results=[
                FileResult(
                    file_path=req.file_path or "",
                    success=True,
                    status="completed",
                    output_path="out",
                    message="ok",
                    error_code=None,
                    details=None,
                )
                for req in batch_request.requests
            ],
            success_count=len(batch_request.requests),
            failed_count=0,
            cancelled=False,
        )

    monkeypatch.setattr(executor.ConversionService, "execute_batch", _stub_execute_batch, raising=True)

    exit_code = executor.execute_batch(
        "convert",
        ["a.docx", "b.docx"],
        options={"x": 1},
        json_mode=True,
        continue_on_error=True,
    )
    out = capsys.readouterr().out

    payload = json.loads(out)
    assert exit_code == 0
    assert payload["success"] is True
    assert payload["command"] == "convert"
    assert payload["data"]["total"] == 2
    assert payload["data"]["processed_count"] == 2
    assert payload["data"]["success_count"] == 2
    assert payload["data"]["failed_count"] == 0


def test_execute_action_fills_error_code_from_docwenerror_in_result(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    e = InvalidInputError("bad", details="d")

    def _stub_execute_action_core(*, action: str, file_path: str, options: dict, progress_callback=None):
        return ConversionResult(success=False, message="generic", error=e, error_code=None), 0.01

    monkeypatch.setattr(executor, "_execute_action_core", _stub_execute_action_core, raising=True)

    exit_code = executor.execute_action("convert", "x.docx", json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 2
    payload = json.loads(out)
    assert payload["success"] is False
    assert payload["command"] == "convert"
    assert payload["data"]["error_code"] == ERROR_CODE_INVALID_INPUT
    assert payload["data"]["details"] == "d"
    assert (payload.get("error") or {}).get("error_code") == ERROR_CODE_INVALID_INPUT


def test_execute_action_keyboard_interrupt_returns_130_json(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def _stub_execute_action_core(*, action: str, file_path: str, options: dict, progress_callback=None):
        raise KeyboardInterrupt()

    monkeypatch.setattr(executor, "_execute_action_core", _stub_execute_action_core, raising=True)

    exit_code = executor.execute_action("convert", "x.docx", json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 130
    payload = json.loads(out)
    assert payload["success"] is False
    assert payload["command"] == "convert"
    assert (payload.get("error") or {}).get("message") == "用户中断"
    assert (payload.get("error") or {}).get("details", {}).get("interrupted") is True


def test_execute_action_keyboard_interrupt_returns_130_text(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def _stub_execute_action_core(*, action: str, file_path: str, options: dict, progress_callback=None):
        raise KeyboardInterrupt()

    monkeypatch.setattr(executor, "_execute_action_core", _stub_execute_action_core, raising=True)

    exit_code = executor.execute_action("convert", "x.docx", json_mode=False)
    captured = capsys.readouterr()

    assert exit_code == 130
    assert "操作已中断" in captured.err


def test_execute_batch_merge_tables_prepends_base_table_and_keeps_result_count(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    captured: dict[str, object] = {}

    def _stub_execute(self, req, **_kwargs):
        captured["file_path"] = req.file_path
        captured["file_list"] = list(req.file_list or [])
        return ConversionResult.ok(output_path="out", message="ok")

    monkeypatch.setattr(executor.ConversionService, "execute", _stub_execute, raising=True)

    exit_code = executor.execute_batch(
        "merge_tables",
        files=["b.xlsx"],
        options={"base_table": "a.xlsx"},
        json_mode=True,
        continue_on_error=True,
    )
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["data"]["total"] == 1
    assert len(payload["data"]["results"]) == 1
    assert captured["file_path"] == "a.xlsx"
    assert captured["file_list"] == ["a.xlsx", "b.xlsx"]
