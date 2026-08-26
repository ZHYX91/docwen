from __future__ import annotations

import importlib
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
CONVERTER = ROOT / "packages/plugins/layout/src/docwen_plugin_layout/to_document/converter.py"


def test_layout_document_priority_fallbacks_match_authoritative_software_config() -> None:
    config = tomllib.loads((ROOT / "configs/software.toml").read_text(encoding="utf-8"))
    converter = importlib.import_module("docwen_plugin_layout.to_document.converter")

    assert list(vars(converter)["_WORD_PROCESSOR_DEFAULT_PRIORITY"]) == config["default_priority"]["word_processors"]
    assert list(vars(converter)["_ODT_DEFAULT_PRIORITY"]) == config["special_conversions"]["odt"]
    assert list(vars(converter)["_PDF_TO_OFFICE_DEFAULT_PRIORITY"]) == config["special_conversions"]["pdf_to_office"]


def test_layout_document_delegates_ordered_execution_and_keeps_odt_wps_illegal() -> None:
    source = CONVERTER.read_text(encoding="utf-8")

    assert "convert_with_backend_priority" in source
    assert "software.default_priority.word_processors" in source
    assert "software.special_conversions.odt" in source
    assert "backend_priority=_document_priority(context, self._target_format)" in source
    assert "backend_priority=_pdf_to_office_priority(context)" in source
    assert 'return {"msoffice_word": candidates["msoffice_word"]}' in source
    assert "convert_with_fallback" not in source
    assert "for backend_id in _pdf_to_office_priority" not in source
