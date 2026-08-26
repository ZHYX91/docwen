from __future__ import annotations

import importlib
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
SMARTDOC = ROOT / "packages/plugins/document/src/docwen_plugin_document/to_document/converter.py"
PRECONVERTER = ROOT / "packages/application/src/docwen_application/preconversion/pre_converter.py"
CONTROLLER = ROOT / "packages/application/src/docwen_application/controller.py"


def test_document_priority_fallbacks_match_authoritative_software_config() -> None:
    config = tomllib.loads((ROOT / "configs/software.toml").read_text(encoding="utf-8"))
    smartdoc = importlib.import_module("docwen_plugin_document.to_document.converter")
    preconverter = importlib.import_module("docwen_application.preconversion.pre_converter")

    assert list(vars(smartdoc)["_DEFAULT_WORD_PRIORITY"]) == config["default_priority"]["word_processors"]
    assert list(vars(smartdoc)["_DEFAULT_ODT_PRIORITY"]) == config["special_conversions"]["odt"]
    assert list(vars(preconverter)["_DEFAULT_WORD_PRIORITY"]) == config["default_priority"]["word_processors"]
    assert list(vars(preconverter)["_DEFAULT_ODT_PRIORITY"]) == config["special_conversions"]["odt"]


def test_smartdoc_and_application_delegate_ordered_execution_to_core() -> None:
    smartdoc = SMARTDOC.read_text(encoding="utf-8")
    preconverter = PRECONVERTER.read_text(encoding="utf-8")
    controller = CONTROLLER.read_text(encoding="utf-8")

    assert "convert_with_backend_priority" in smartdoc
    assert "software.default_priority.word_processors" in smartdoc
    assert "software.special_conversions.odt" in smartdoc
    assert "convert_with_fallback" not in smartdoc
    assert "convert_with_backend_priority" in preconverter
    assert "backend_priority=selected_priority" in preconverter
    assert 'com_candidates.pop("wps_writer")' in preconverter
    assert "convert_with_fallback" not in preconverter
    assert "captured_snapshot = self._config_port.snapshot()" in controller
    assert "self._config_port.get(key, list(default))" not in controller
    assert "backend_priority=backend_priority" in controller
