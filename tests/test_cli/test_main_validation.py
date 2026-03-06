"""CLI 单元测试。"""

from __future__ import annotations

import json
import importlib
from types import SimpleNamespace
from pathlib import Path

import pytest

cli_main = importlib.import_module("docwen.cli.main")

pytestmark = pytest.mark.unit


def _base_args(**overrides):
    args = SimpleNamespace(
        files=["a.docx"],
        json=True,
        quiet=False,
        verbose=False,
        continue_on_error=True,
        jobs=1,
        batch=False,
        timing=False,
        to="md",
        ocr=False,
        extract_img=False,
        no_extract_img=False,
        optimize_for=None,
        template=None,
        check=None,
        mode=None,
        base_table=None,
        pages=None,
        dpi=None,
        compress=None,
        size_limit=None,
        quality_mode=None,
        clean_numbering=None,
        add_numbering=None,
    )
    for k, v in overrides.items():
        setattr(args, k, v)
    return args


def test_build_options_does_not_force_extract_img_when_ocr_is_enabled() -> None:
    args = _base_args(ocr=True, extract_img=False, no_extract_img=False)
    options = cli_main.build_options("convert", args)
    assert options["extract_ocr"] is True
    assert "extract_image" not in options


def test_build_options_allows_ocr_when_extract_img_is_disabled() -> None:
    args = _base_args(ocr=True, extract_img=False, no_extract_img=True)
    options = cli_main.build_options("convert", args)
    assert options["extract_ocr"] is True
    assert options["extract_image"] is False


def test_build_options_rejects_check_none_with_other_checks() -> None:
    args = _base_args(check=["none", "punct"])
    with pytest.raises(ValueError, match="--check none 不能与其它 --check 同时使用"):
        cli_main.build_options("validate", args)


def test_build_options_convert_numbering_keys_are_canonical() -> None:
    args = _base_args(clean_numbering="remove", add_numbering="gongwen_standard")
    options = cli_main.build_options("convert", args)
    assert options["clean_numbering"] == "remove"
    assert options["add_numbering_mode"] == "gongwen_standard"
    assert "doc_remove_numbering" not in options
    assert "md_remove_numbering" not in options
    assert "remove_numbering" not in options
    assert "numbering_scheme" not in options


def test_build_options_convert_numbering_modes_absent_when_not_provided() -> None:
    args = _base_args(clean_numbering=None, add_numbering=None)
    options = cli_main.build_options("convert", args)
    assert "clean_numbering" not in options
    assert "add_numbering_mode" not in options


def test_execute_headless_json_invalid_files_warning_goes_to_stderr(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    args = _base_args(json=True, quiet=False)

    monkeypatch.setattr(cli_main, "cli_t", lambda _key, **_kwargs: "WARN", raising=True)
    monkeypatch.setattr("docwen.cli.utils.expand_paths", lambda files: list(files), raising=True)
    monkeypatch.setattr(
        "docwen.cli.utils.validate_files",
        lambda _files: (["ok.docx"], [("bad.bin", "unsupported")], []),
        raising=True,
    )
    monkeypatch.setattr("docwen.cli.utils.create_progress_callback", lambda **_kwargs: (lambda _msg: None), raising=True)

    def _stub_execute_action(_action, _file, _options, *, json_mode, **_kwargs):
        assert json_mode is True
        print(json.dumps({"ok": True}))
        return 0

    monkeypatch.setattr("docwen.cli.executor.execute_action", _stub_execute_action, raising=True)

    exit_code = cli_main._run_action_for_files("convert", args.files, options={}, args=args)
    captured = capsys.readouterr()

    assert exit_code == 0
    assert json.loads(captured.out) == {"ok": True}
    assert "WARN" in captured.err
    assert "bad.bin" in captured.err


def _run_main_with_args(monkeypatch: pytest.MonkeyPatch, args: SimpleNamespace) -> int:
    class _Parser:
        def parse_args(self):
            return args

    monkeypatch.setattr(cli_main, "setup_console_encoding", lambda: None, raising=True)
    monkeypatch.setattr(cli_main, "pre_parse_lang", lambda: None, raising=True)
    monkeypatch.setattr(cli_main, "init_cli_locale", lambda _lang=None: None, raising=True)
    monkeypatch.setattr("docwen.translation.set_translator", lambda _t: None, raising=True)
    monkeypatch.setattr(cli_main, "create_argument_parser", lambda: _Parser(), raising=True)
    monkeypatch.setattr(cli_main, "cli_t", lambda key, **_kwargs: "错误" if key == "cli.messages.error_prefix" else key, raising=True)
    return int(cli_main.main())


def test_convert_md_to_docx_requires_template_json(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    f = tmp_path / "a.md"
    f.write_text("# t\n", encoding="utf-8")

    args = _base_args(files=[str(f)], to="docx", template=None, json=True)
    args.command = "convert"

    monkeypatch.setattr("docwen.cli.executor.get_supported_convert_targets", lambda: ["md", "docx"], raising=True)

    exit_code = _run_main_with_args(monkeypatch, args)
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == 2
    assert payload["success"] is False
    assert payload["command"] == "convert"
    assert payload["error"]["error_code"] == "invalid_input"
    assert "需要指定 --template" in payload["error"]["message"]


def test_convert_to_non_md_rejects_md_only_flags_json(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], tmp_path: Path
) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    args = _base_args(files=[str(f)], to="docx", extract_img=True, json=True)
    args.command = "convert"

    monkeypatch.setattr("docwen.cli.executor.get_supported_convert_targets", lambda: ["md", "docx"], raising=True)

    exit_code = _run_main_with_args(monkeypatch, args)
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == 2
    assert payload["success"] is False
    assert payload["command"] == "convert"
    assert payload["error"]["error_code"] == "invalid_input"
    assert "仅在 --to md 时可用" in payload["error"]["message"]


def test_convert_to_md_rejects_unknown_optimize_for_json(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], tmp_path: Path
) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    args = _base_args(files=[str(f)], to="md", optimize_for="gongwne", json=True)
    args.command = "convert"

    monkeypatch.setattr("docwen.cli.executor.get_supported_convert_targets", lambda: ["md", "docx"], raising=True)
    monkeypatch.setattr(
        "docwen.config.config_manager.get_optimization_types",
        lambda: {"gongwen": {"enabled": True}},
        raising=True,
    )

    exit_code = _run_main_with_args(monkeypatch, args)
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == 2
    assert payload["success"] is False
    assert payload["command"] == "convert"
    assert payload["error"]["error_code"] == "invalid_input"
    assert "--optimize-for" in payload["error"]["message"]
    assert "你可能想要" in payload["error"]["message"]


def test_convert_to_md_allows_ocr_without_extract_img_passes_options_to_executor(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")

    args = _base_args(files=[str(f)], to="md", ocr=True, no_extract_img=True, json=True)
    args.command = "convert"

    monkeypatch.setattr("docwen.cli.executor.get_supported_convert_targets", lambda: ["md"], raising=True)

    captured = {}

    def _stub_run_action_for_files(action: str, raw_files: list[str], options: dict, args):
        captured["action"] = action
        captured["raw_files"] = list(raw_files)
        captured["options"] = dict(options)
        return 0

    monkeypatch.setattr(cli_main, "_run_action_for_files", _stub_run_action_for_files, raising=True)

    exit_code = _run_main_with_args(monkeypatch, args)
    assert exit_code == 0
    assert captured["action"] == "convert"
    assert captured["raw_files"] == [str(f)]
    assert captured["options"]["target_format"] == "md"
    assert captured["options"]["extract_ocr"] is True
    assert captured["options"]["extract_image"] is False
