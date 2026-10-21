"""PNG precision warnings use the GUI locale without changing machine facts."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

from docwen_core.models.result import ConversionDiagnostic, ConversionResult
from docwen_gui.execution_presenter import _result_warning_messages
from docwen_gui.i18n import get_locale, set_locale

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[4]
LOCALES = ("ar_SA", "de_DE", "en_US", "es_ES", "fr_FR", "ja_JP", "ko_KR", "pt_BR", "ru_RU", "vi_VN", "zh_CN", "zh_TW")
RAW_MESSAGE = "Unchanged machine diagnostic text."


@pytest.mark.parametrize("locale", LOCALES)
@pytest.mark.parametrize("location", ["", "integer.tif:frame-2"])
def test_png_clipping_warning_uses_every_gui_locale_and_preserves_location(locale: str, location: str) -> None:
    table = tomllib.loads((ROOT / "i18n" / "locales" / f"{locale}.toml").read_text(encoding="utf-8"))
    expected = table["main_window"]["image_png_16bit_clipped"]
    assert isinstance(expected, str) and expected.strip()
    assert "PNG" in expected and "TIFF" in expected and "0–65535" in expected
    diagnostic = ConversionDiagnostic("warning", RAW_MESSAGE, code="IMAGE-PNG-16BIT-CLIPPED", location=location)
    result = ConversionResult(task_id="precision", success=True, diagnostics=[diagnostic])
    previous = get_locale()
    try:
        set_locale(locale)
        assert _result_warning_messages(result) == [f"{expected} ({location})" if location else expected]
    finally:
        set_locale(previous)
    assert diagnostic.message == RAW_MESSAGE
    assert diagnostic.code == "IMAGE-PNG-16BIT-CLIPPED"
    assert diagnostic.location == location


@pytest.mark.parametrize("locale,phrase", [("zh_CN", "超出"), ("en_US", "outside")])
def test_png_precision_translation_is_independent_of_machine_message(locale: str, phrase: str) -> None:
    previous = get_locale()
    try:
        set_locale(locale)
        result = ConversionResult(
            task_id="precision",
            success=True,
            diagnostics=[ConversionDiagnostic("warning", "", code="IMAGE-PNG-16BIT-CLIPPED")],
        )
        assert phrase in _result_warning_messages(result)[0]
    finally:
        set_locale(previous)


@pytest.mark.parametrize("locale", ["zh_CN", "en_US"])
def test_unknown_warning_keeps_existing_raw_fallback(locale: str) -> None:
    previous = get_locale()
    try:
        set_locale(locale)
        result = ConversionResult(
            task_id="other",
            success=True,
            diagnostics=[
                ConversionDiagnostic("warning", RAW_MESSAGE, code="OTHER-WARNING", location="page-3"),
                ConversionDiagnostic("info", "not a warning", code="IMAGE-PNG-16BIT-CLIPPED"),
            ],
        )
        assert _result_warning_messages(result) == [f"{RAW_MESSAGE} (page-3)"]
    finally:
        set_locale(previous)
