"""Regression guards for reviewed About-page product positioning."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

_LOCALES = Path(__file__).resolve().parents[2] / "i18n" / "locales"

_EXPECTED = {
    "zh_CN": "文档与表格格式转换工具（公文转换器）",
    "zh_TW": "文件與表格格式轉換工具（公文轉換器）",
    "de_DE": "Werkzeug zur Konvertierung von Dokument- und Tabellenformaten",
    "es_ES": "Herramienta de conversión de documentos y hojas de cálculo (convertidor de documentos oficiales)",
    "fr_FR": "Outil de conversion de documents et de feuilles de calcul (Convertisseur de documents officiels)",
    "ja_JP": "文書・表形式変換ツール（公文書コンバーター）",
    "ko_KR": "문서 및 스프레드시트 형식 변환 도구(공문 변환기)",
}


@pytest.mark.parametrize(("locale", "expected"), sorted(_EXPECTED.items()))
def test_about_subtitle_uses_document_table_positioning(locale: str, expected: str) -> None:
    data = tomllib.loads((_LOCALES / f"{locale}.toml").read_text(encoding="utf-8"))
    assert data["about"]["subtitle"] == expected
