"""GUI 逻辑单元测试。"""

from __future__ import annotations

import dataclasses
import pytest

from docwen.services.context import get_default_context
from docwen.services.requests import ConversionRequest
from docwen.services.result import ConversionResult
from docwen.services.use_cases import ConversionService

pytestmark = [pytest.mark.unit, pytest.mark.windows_only]


def test_service_merges_proofread_defaults_when_missing_keys(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    captured = {}

    class _CM:
        def get_proofread_engine_config(self):
            return {
                "enable_symbol_pairing": True,
                "enable_symbol_correction": False,
                "enable_typos_rule": True,
                "enable_sensitive_word": False,
            }

    def _fake_get_strategy(*, action_type=None, source_format=None, target_format=None):
        class _Strategy:
            def execute(self, file_path: str, options=None, progress_callback=None):
                captured["options"] = dict(options or {})
                return ConversionResult.ok(message="ok", output_path=file_path)

        return _Strategy

    ctx = dataclasses.replace(get_default_context(), config_manager=_CM(), get_strategy=_fake_get_strategy)

    f = tmp_path / "a.md"
    f.write_text("# hi", encoding="utf-8")
    result = ConversionService(ctx=ctx).execute(ConversionRequest(file_path=str(f), action_type="convert_md_to_docx"))
    assert result.success is True
    assert captured["options"]["proofread_options"] == {
        "symbol_pairing": True,
        "symbol_correction": False,
        "typos_rule": True,
        "sensitive_word": False,
    }


def test_service_proofread_overrides_config_defaults_when_keys_present(
    monkeypatch: pytest.MonkeyPatch, tmp_path
) -> None:
    captured = {}

    class _CM:
        def get_proofread_engine_config(self):
            return {
                "enable_symbol_pairing": True,
                "enable_symbol_correction": True,
                "enable_typos_rule": True,
                "enable_sensitive_word": True,
            }

    def _fake_get_strategy(*, action_type=None, source_format=None, target_format=None):
        class _Strategy:
            def execute(self, file_path: str, options=None, progress_callback=None):
                captured["options"] = dict(options or {})
                return ConversionResult.ok(message="ok", output_path=file_path)

        return _Strategy

    ctx = dataclasses.replace(get_default_context(), config_manager=_CM(), get_strategy=_fake_get_strategy)

    f = tmp_path / "a.md"
    f.write_text("# hi", encoding="utf-8")
    result = ConversionService(ctx=ctx).execute(
        ConversionRequest(file_path=str(f), action_type="convert_md_to_docx", options={"typos_rule": False})
    )
    assert result.success is True
    assert captured["options"]["proofread_options"] == {
        "symbol_pairing": True,
        "symbol_correction": True,
        "typos_rule": False,
        "sensitive_word": True,
    }


def test_service_proofread_all_false_is_preserved(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    captured = {}

    def _fake_get_strategy(*, action_type=None, source_format=None, target_format=None):
        class _Strategy:
            def execute(self, file_path: str, options=None, progress_callback=None):
                captured["options"] = dict(options or {})
                return ConversionResult.ok(message="ok", output_path=file_path)

        return _Strategy

    ctx = dataclasses.replace(get_default_context(), get_strategy=_fake_get_strategy)

    f = tmp_path / "a.md"
    f.write_text("# hi", encoding="utf-8")
    result = ConversionService(ctx=ctx).execute(
        ConversionRequest(
            file_path=str(f),
            action_type="convert_md_to_docx",
            options={
                "symbol_pairing": False,
                "symbol_correction": False,
                "typos_rule": False,
                "sensitive_word": False,
            },
        )
    )
    assert result.success is True
    assert captured["options"]["proofread_options"] == {
        "symbol_pairing": False,
        "symbol_correction": False,
        "typos_rule": False,
        "sensitive_word": False,
    }
