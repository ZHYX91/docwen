from __future__ import annotations

import json
from pathlib import Path

import pytest

import docwen.cli.executor as executor
from docwen.services.result import ConversionResult


pytestmark = pytest.mark.unit


class _OkStrategy:
    def execute(self, *, file_path: str, options: dict, progress_callback=None):
        return ConversionResult(success=True, output_path="out", message="ok")


class _FailStrategy:
    def execute(self, *, file_path: str, options: dict, progress_callback=None):
        raise RuntimeError("boom")


def test_get_strategy_for_action_export_md_uses_category(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
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

    assert executor.get_strategy_for_action("export_md", str(f), {"actual_format": "docx"}) is _OkStrategy
    assert called == {"action_type": None, "source_format": "document", "target_format": "md"}


def test_get_strategy_for_action_convert_requires_target(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(executor, "get_strategy", lambda **_kwargs: _OkStrategy, raising=True)

    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    with pytest.raises(ValueError) as excinfo:
        executor.get_strategy_for_action("convert", str(f), {"actual_format": "docx"})

    assert "需要指定目标格式" in str(excinfo.value)


def test_execute_action_missing_file_returns_1_text(capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    missing = tmp_path / "missing.docx"
    exit_code = executor.execute_action("export_md", str(missing), json_mode=False)
    out = capsys.readouterr().out

    assert exit_code == 1
    assert "文件不存在" in out


def test_execute_action_missing_file_returns_1_json(capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    missing = tmp_path / "missing.docx"
    exit_code = executor.execute_action("export_md", str(missing), json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 1
    payload = json.loads(out)
    assert payload["success"] is False
    assert payload["action"] == "export_md"
    assert "文件不存在" in payload["error"]


def test_execute_action_actual_format_detection_failure_is_non_fatal(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    monkeypatch.setattr(executor, "get_strategy_for_action", lambda *_args, **_kwargs: _OkStrategy, raising=True)
    monkeypatch.setattr(
        "docwen.utils.file_type_utils.detect_actual_file_format",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("x")),
        raising=True,
    )

    assert executor.execute_action("export_md", str(f), json_mode=True) == 0


def test_execute_action_strategy_exception_returns_1_json(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    monkeypatch.setattr(executor, "get_strategy_for_action", lambda *_args, **_kwargs: _FailStrategy, raising=True)

    exit_code = executor.execute_action("export_md", str(f), json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 1
    payload = json.loads(out)
    assert payload["success"] is False
    assert payload["action"] == "export_md"
    assert "boom" in payload["error"]


def test_get_supported_actions_covers_categories() -> None:
    markdown = executor.get_supported_actions("markdown", "md")
    assert {a["name"] for a in markdown} == {"convert_md_to_docx", "convert_md_to_xlsx"}

    document = executor.get_supported_actions("document", "docx")
    assert {a["name"] for a in document} == {"export_md", "convert", "validate"}

    spreadsheet = executor.get_supported_actions("spreadsheet", "xlsx")
    assert {a["name"] for a in spreadsheet} == {"export_md", "convert", "merge_tables"}

    layout = executor.get_supported_actions("layout", "pdf")
    assert {a["name"] for a in layout} == {"export_md", "merge_pdfs", "split_pdf"}

    image = executor.get_supported_actions("image", "png")
    assert {a["name"] for a in image} == {"export_md", "convert"}

    assert executor.get_supported_actions("unknown", "bin") == []


def test_inspect_file_json(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    monkeypatch.setattr("docwen.cli.utils.detect_category", lambda *_args, **_kwargs: "markdown", raising=True)
    monkeypatch.setattr("docwen.cli.utils.detect_format", lambda *_args, **_kwargs: "md", raising=True)

    exit_code = executor.inspect_file("x.md", json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert payload["category"] == "markdown"
    assert payload["format"] == "md"
    assert {a["name"] for a in payload["supported_actions"]} == {"convert_md_to_docx", "convert_md_to_xlsx"}


def test_list_all_actions_json(capsys: pytest.CaptureFixture[str]) -> None:
    exit_code = executor.list_all_actions(json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert "actions" in payload
    assert any(a["name"] == "export_md" for a in payload["actions"])


def test_list_numbering_schemes_fallback_json(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    import importlib

    cm = importlib.import_module("docwen.config.config_manager")
    monkeypatch.setattr(cm.config_manager, "get_heading_schemes", lambda: (_ for _ in ()).throw(RuntimeError("x")), raising=True)

    exit_code = executor.list_numbering_schemes(json_mode=True)
    out = capsys.readouterr().out

    assert exit_code == 0
    payload = json.loads(out)
    assert "schemes" in payload
    assert any(s["id"] == "gongwen_standard" for s in payload["schemes"])


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
    assert payload["action"] == "list_templates"
    assert "no templates" in payload["error"]
