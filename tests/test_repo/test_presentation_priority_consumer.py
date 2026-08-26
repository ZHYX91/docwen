from __future__ import annotations

import importlib
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
CONVERTER = ROOT / "packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/ppt_converter.py"


def test_presentation_priority_fallback_matches_authoritative_software_config() -> None:
    config = tomllib.loads((ROOT / "configs/software.toml").read_text(encoding="utf-8"))
    converter = importlib.import_module("docwen_plugin_presentation.pptx_md.ppt_converter")

    assert (
        list(vars(converter)["_DEFAULT_PRESENTATION_PRIORITY"]) == config["default_priority"]["presentation_processors"]
    )


def test_presentation_delegates_configured_ordered_execution_to_core() -> None:
    source = CONVERTER.read_text(encoding="utf-8")

    assert "convert_with_backend_priority" in source
    assert "software.default_priority.presentation_processors" in source
    assert '"wps_presentation": BridgeCandidate(' in source
    assert '"msoffice_powerpoint": BridgeCandidate(' in source
    assert "backend_priority=_configured_priority(context)" in source
    assert "convert_with_fallback" not in source
