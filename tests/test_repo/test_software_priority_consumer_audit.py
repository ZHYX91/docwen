from __future__ import annotations

import importlib
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
MARKDOWN_CONVERTER = ROOT / "packages/plugins/markdown/src/docwen_plugin_markdown/office_bridge/converter.py"
PRINT_CONVERTER = ROOT / "packages/plugins/print/src/docwen_plugin_print/paged_output/converter.py"


def test_markdown_priority_fallbacks_match_authoritative_software_config() -> None:
    config = tomllib.loads((ROOT / "configs/software.toml").read_text(encoding="utf-8"))
    module = importlib.import_module("docwen_plugin_markdown.office_bridge.converter")
    defaults = vars(module)["_DEFAULT_PRIORITIES"]
    keys = vars(module)["_PRIORITY_KEYS"]

    assert list(defaults["word_processors"]) == config["default_priority"]["word_processors"]
    assert list(defaults["spreadsheet_processors"]) == config["default_priority"]["spreadsheet_processors"]
    assert list(defaults["odt"]) == config["special_conversions"]["odt"]
    assert list(defaults["ods"]) == config["special_conversions"]["ods"]
    assert list(defaults["document_to_pdf"]) == config["special_conversions"]["document_to_pdf"]
    assert keys == {
        "word_processors": "software.default_priority.word_processors",
        "spreadsheet_processors": "software.default_priority.spreadsheet_processors",
        "odt": "software.special_conversions.odt",
        "ods": "software.special_conversions.ods",
        "document_to_pdf": "software.special_conversions.document_to_pdf",
    }


def test_markdown_priority_consumer_delegates_ordered_execution_to_core() -> None:
    source = MARKDOWN_CONVERTER.read_text(encoding="utf-8")

    assert "from docwen_core.office_bridge import BridgeCandidate, convert_with_backend_priority" in source
    assert "bridge_result = convert_with_backend_priority(" in source
    assert "backend_priority=self._configured_priority(context, priority_category)" in source
    assert "software.special_conversions.document_to_pdf" in source
    assert 'candidates.pop("wps_writer")' in source
    assert 'candidates.pop("wps_spreadsheets")' in source
    assert "convert_with_fallback" not in source


def test_print_priority_consumer_shares_the_core_ordered_executor() -> None:
    source = PRINT_CONVERTER.read_text(encoding="utf-8")

    assert "convert_with_backend_priority" in source
    assert "backend_priority=_configured_priority(context, category)" in source
    assert "com_candidates=_COM_CANDIDATES[category]" in source
    assert "convert_with_fallback" not in source
    assert "for backend_id in _configured_priority" not in source
