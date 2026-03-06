"""services 单元测试。"""

from __future__ import annotations

import dataclasses
from typing import Any

import pytest

from docwen.proofread_keys import PROOFREAD_OPTION_KEYS, SYMBOL_PAIRING
from docwen.services.context import get_default_context
from docwen.services.error_codes import ERROR_CODE_INVALID_INPUT
from docwen.services.requests import ConversionRequest
from docwen.services.result import ConversionResult
from docwen.services.use_cases import ConversionService

pytestmark = pytest.mark.unit


class _Capture:
    def __init__(self) -> None:
        self.called_with: list[tuple[str | None, str | None, str | None]] = []
        self.last_options: dict[str, Any] | None = None


def test_named_action_rejects_source_target(tmp_path, monkeypatch):
    capture = _Capture()

    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        capture.called_with.append((action_type, source_format, target_format))

        class S:
            def execute(self, file_path: str, options=None, progress_callback=None):
                return ConversionResult.ok(output_path=file_path)

        return S

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    p = tmp_path / "a.md"
    p.write_text("# hi", encoding="utf-8")

    req = ConversionRequest(
        file_path=str(p),
        action_type="validate",
        target_format="md",
    )
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is False
    assert result.error_code == ERROR_CODE_INVALID_INPUT


def test_conversion_requires_target_format(tmp_path, monkeypatch):
    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        raise AssertionError("should not be called")

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    p = tmp_path / "a.md"
    p.write_text("# hi", encoding="utf-8")

    req = ConversionRequest(file_path=str(p))
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is False
    assert result.error_code == ERROR_CODE_INVALID_INPUT


def test_convert_action_type_is_normalized_to_source_target(tmp_path, monkeypatch):
    capture = _Capture()

    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        capture.called_with.append((action_type, source_format, target_format))

        class S:
            def execute(self, file_path: str, options=None, progress_callback=None):
                capture.last_options = dict(options or {})
                return ConversionResult.ok(output_path=file_path)

        return S

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    p = tmp_path / "a.md"
    p.write_text("# hi", encoding="utf-8")

    req = ConversionRequest(file_path=str(p), action_type="convert_md_to_docx")
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is True
    assert capture.called_with[-1] == (None, "md", "docx")


def test_convert_to_md_uses_category_as_source_format(tmp_path, monkeypatch):
    capture = _Capture()

    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        capture.called_with.append((action_type, source_format, target_format))

        class S:
            def execute(self, file_path: str, options=None, progress_callback=None):
                capture.last_options = dict(options or {})
                return ConversionResult.ok(output_path=file_path)

        return S

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    p = tmp_path / "a.docx"
    p.write_text("x", encoding="utf-8")

    req = ConversionRequest(file_path=str(p), target_format="md", actual_format="docx", category="document")
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is True
    assert capture.called_with[-1] == (None, "document", "md")
    assert capture.last_options is not None
    assert capture.last_options["extract_image"] is True
    assert capture.last_options["extract_ocr"] is False
    assert capture.last_options["optimize_for_type"] == ""


def test_convert_to_md_defaults_depend_on_category(tmp_path, monkeypatch):
    capture = _Capture()

    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        capture.called_with.append((action_type, source_format, target_format))

        class S:
            def execute(self, file_path: str, options=None, progress_callback=None):
                capture.last_options = dict(options or {})
                return ConversionResult.ok(output_path=file_path)

        return S

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    p = tmp_path / "a.pdf"
    p.write_bytes(b"%PDF-1.4")

    req = ConversionRequest(file_path=str(p), target_format="md", actual_format="pdf", category="layout")
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is True
    assert capture.called_with[-1] == (None, "layout", "md")
    assert capture.last_options is not None
    assert capture.last_options["extract_image"] is False
    assert capture.last_options["extract_ocr"] is False
    assert capture.last_options["optimize_for_type"] == ""


def test_proofread_options_flat_keys_are_merged_with_defaults(tmp_path, monkeypatch):
    capture = _Capture()

    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        class S:
            def execute(self, file_path: str, options=None, progress_callback=None):
                capture.last_options = dict(options or {})
                return ConversionResult.ok(output_path=file_path)

        return S

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    p = tmp_path / "a.md"
    p.write_text("# hi", encoding="utf-8")

    req = ConversionRequest(
        file_path=str(p),
        action_type="convert_md_to_docx",
        options={SYMBOL_PAIRING: False},
    )
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is True
    assert capture.last_options is not None
    proofread = capture.last_options.get("proofread_options")
    assert isinstance(proofread, dict)
    assert set(proofread.keys()) == set(PROOFREAD_OPTION_KEYS)
    assert proofread[SYMBOL_PAIRING] is False


def test_proofread_options_empty_dict_means_disable_all(tmp_path, monkeypatch):
    capture = _Capture()

    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        class S:
            def execute(self, file_path: str, options=None, progress_callback=None):
                capture.last_options = dict(options or {})
                return ConversionResult.ok(output_path=file_path)

        return S

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    p = tmp_path / "a.md"
    p.write_text("# hi", encoding="utf-8")

    req = ConversionRequest(
        file_path=str(p),
        action_type="convert_md_to_docx",
        options={"proofread_options": {}},
    )
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is True
    assert capture.last_options is not None
    assert capture.last_options.get("proofread_options") == {}


def test_merge_tables_requires_file_list_contains_base(tmp_path, monkeypatch):
    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        raise AssertionError("should not be called")

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    base = tmp_path / "base.xlsx"
    base.write_bytes(b"")
    other = tmp_path / "other.xlsx"
    other.write_bytes(b"")

    req = ConversionRequest(
        file_path=str(base),
        action_type="merge_tables",
        file_list=[str(other), str(other)],
        options={"mode": "row"},
        actual_format="xlsx",
        category="spreadsheet",
    )
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is False
    assert result.error_code == ERROR_CODE_INVALID_INPUT


def test_merge_tables_mode_is_mapped(tmp_path, monkeypatch):
    capture = _Capture()

    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        class S:
            def execute(self, file_path: str, options=None, progress_callback=None):
                capture.last_options = dict(options or {})
                return ConversionResult.ok(output_path=file_path)

        return S

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    base = tmp_path / "base.xlsx"
    base.write_bytes(b"")
    other = tmp_path / "other.xlsx"
    other.write_bytes(b"")

    req = ConversionRequest(
        file_path=str(base),
        action_type="merge_tables",
        file_list=[str(base), str(other)],
        options={"mode": "row"},
        actual_format="xlsx",
        category="spreadsheet",
    )
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is True
    assert capture.last_options is not None
    assert capture.last_options["mode"] == 1


def test_merge_tables_mode_invalid_values_fail(tmp_path, monkeypatch):
    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        raise AssertionError("should not be called")

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    base = tmp_path / "base.xlsx"
    base.write_bytes(b"")
    other = tmp_path / "other.xlsx"
    other.write_bytes(b"")

    for mode in (0, "bad", None):
        req = ConversionRequest(
            file_path=str(base),
            action_type="merge_tables",
            file_list=[str(base), str(other)],
            options={"mode": mode},
            actual_format="xlsx",
            category="spreadsheet",
        )
        result = ConversionService(ctx=ctx).execute(req)
        assert result.success is False
        assert result.error_code == ERROR_CODE_INVALID_INPUT


def test_image_same_format_with_size_limit_is_not_skipped(tmp_path, monkeypatch):
    capture = _Capture()

    def fake_get_strategy(action_type=None, source_format=None, target_format=None):
        capture.called_with.append((action_type, source_format, target_format))

        class S:
            def execute(self, file_path: str, options=None, progress_callback=None):
                capture.last_options = dict(options or {})
                return ConversionResult.ok(output_path=file_path)

        return S

    ctx = dataclasses.replace(get_default_context(), get_strategy=fake_get_strategy)

    p = tmp_path / "a.png"
    p.write_bytes(b"\x89PNG\r\n\x1a\n")

    req = ConversionRequest(
        file_path=str(p),
        target_format="png",
        actual_format="png",
        category="image",
        options={"compress_mode": "limit_size", "size_limit": 50},
    )
    result = ConversionService(ctx=ctx).execute(req)
    assert result.success is True
    assert capture.called_with
