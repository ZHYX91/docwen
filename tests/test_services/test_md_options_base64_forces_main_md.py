"""services 单元测试。"""

from __future__ import annotations

from typing import Any

import pytest

from docwen.services.context import AppContext
from docwen.services.use_cases import _normalize_md_options

pytestmark = pytest.mark.unit


class _DummyConfigManager:
    def __init__(self, extraction_mode: str, ocr_placement_mode: str) -> None:
        self._extraction_mode = extraction_mode
        self._ocr_placement_mode = ocr_placement_mode

    def get_export_to_md_image_extraction_mode(self) -> str:
        return self._extraction_mode

    def get_export_to_md_ocr_placement_mode(self) -> str:
        return self._ocr_placement_mode


def _make_ctx(extraction_mode: str, ocr_placement_mode: str) -> AppContext:
    cfg = _DummyConfigManager(extraction_mode=extraction_mode, ocr_placement_mode=ocr_placement_mode)
    return AppContext(
        t=lambda *_a, **_k: "",
        config_manager=cfg,
        detect_actual_file_format=lambda _p: "unknown",
        get_actual_file_category=lambda _p, _f=None: "unknown",
        get_strategy=lambda **_k: None,
    )


def test_normalize_md_options_forces_main_md_when_base64_provided() -> None:
    options: dict[str, Any] = {
        "extract_image": True,
        "extract_ocr": True,
        "to_md_image_extraction_mode": "base64",
        "to_md_ocr_placement_mode": "image_md",
    }
    _normalize_md_options(options, category="document", ctx=_make_ctx("file", "image_md"))
    assert options["to_md_image_extraction_mode"] == "base64"
    assert options["to_md_ocr_placement_mode"] == "main_md"


def test_normalize_md_options_forces_main_md_when_base64_from_config() -> None:
    options: dict[str, Any] = {"extract_image": True, "extract_ocr": True}
    _normalize_md_options(options, category="document", ctx=_make_ctx("base64", "image_md"))
    assert options["to_md_image_extraction_mode"] == "base64"
    assert options["to_md_ocr_placement_mode"] == "main_md"
