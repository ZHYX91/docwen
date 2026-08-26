from __future__ import annotations

import importlib
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
CONVERTER = ROOT / "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/format_conversion/converter.py"


def test_smartsheet_priority_fallbacks_match_authoritative_software_config() -> None:
    config = tomllib.loads((ROOT / "configs/software.toml").read_text(encoding="utf-8"))
    converter = importlib.import_module("docwen_plugin_spreadsheet.format_conversion.converter")

    assert (
        list(vars(converter)["_DEFAULT_SPREADSHEET_PRIORITY"]) == config["default_priority"]["spreadsheet_processors"]
    )
    assert list(vars(converter)["_DEFAULT_ODS_PRIORITY"]) == config["special_conversions"]["ods"]


def test_smartsheet_delegates_per_leg_ordered_execution_to_core() -> None:
    source = CONVERTER.read_text(encoding="utf-8")

    assert "convert_with_backend_priority" in source
    assert "software.default_priority.spreadsheet_processors" in source
    assert "software.special_conversions.ods" in source
    assert 'return {"msoffice_excel": candidates["msoffice_excel"]}' in source
    assert "backend_priority=_configured_priority(" in source
    assert "convert_with_fallback" not in source
