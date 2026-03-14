"""utils 单元测试。"""

from __future__ import annotations

import sys
import threading
import time
import types
from pathlib import Path

import pytest

import docwen.utils.ocr_utils as ocr_utils

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_resolve_ocr_language_auto_by_locale(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ocr_utils, "get_configured_ocr_language", lambda: ocr_utils.OCR_LANGUAGE_AUTO, raising=True)
    monkeypatch.setattr(ocr_utils, "LOCALE_TO_OCR_LANGUAGE", {"en_US": ocr_utils.OCR_LANGUAGE_ENGLISH}, raising=True)

    import docwen.i18n as i18n

    monkeypatch.setattr(i18n, "get_current_locale", lambda: "en_US", raising=True)
    assert ocr_utils.resolve_ocr_language() == ocr_utils.OCR_LANGUAGE_ENGLISH


@pytest.mark.unit
def test_get_configured_ocr_language_uses_config_manager(monkeypatch: pytest.MonkeyPatch) -> None:
    import importlib

    cm = importlib.import_module("docwen.config.config_manager")

    monkeypatch.setattr(cm.config_manager, "get_ocr_language", lambda: ocr_utils.OCR_LANGUAGE_LATIN, raising=True)
    assert ocr_utils.get_configured_ocr_language() == ocr_utils.OCR_LANGUAGE_LATIN


@pytest.mark.unit
def test_get_ocr_thread_safe_single_init(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    ocr_utils.reset_ocr()

    model_dir = tmp_path / "models" / "rapidocr"
    model_dir.mkdir(parents=True)

    models = ocr_utils.OCR_LANGUAGE_MODELS[ocr_utils.OCR_LANGUAGE_CHINESE]
    for name in models.values():
        (model_dir / name).write_bytes(b"x")

    monkeypatch.setattr("docwen.utils.path_utils.get_project_root", lambda: str(tmp_path), raising=True)

    init_count = {"n": 0}

    class StubRapidOCR:
        def __init__(self, **_kwargs):
            init_count["n"] += 1
            time.sleep(0.02)

        def __call__(self, _path: str):
            return [], []

    stub_module = types.ModuleType("rapidocr_onnxruntime")
    stub_module.RapidOCR = StubRapidOCR
    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", stub_module)

    errors: list[Exception] = []

    def _worker() -> None:
        try:
            ocr_utils.get_ocr(ocr_language=ocr_utils.OCR_LANGUAGE_CHINESE)
        except Exception as e:
            errors.append(e)

    threads = [threading.Thread(target=_worker) for _ in range(10)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()

    assert errors == []
    assert init_count["n"] == 1


@pytest.mark.unit
def test_estimate_font_size_nearest_common() -> None:
    assert ocr_utils.estimate_font_size(16, image_dpi=96) == 12


@pytest.mark.unit
def test_extract_text_simple_cancelled() -> None:
    ev = threading.Event()
    ev.set()
    assert ocr_utils.extract_text_simple("any.png", cancel_event=ev) == ""


@pytest.mark.unit
def test_extract_text_simple_with_stub_ocr(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    image_path = tmp_path / "x.png"
    image_path.write_bytes(b"x")

    class StubOCR:
        def __call__(self, _path: str):
            return [
                [[[0, 0], [10, 0], [10, 10], [0, 10]], "A", 0.9],
                [[[0, 0], [10, 0], [10, 10], [0, 10]], "B", 0.4],
            ], [0.01]

    monkeypatch.setattr(ocr_utils, "get_ocr", lambda *_args, **_kwargs: StubOCR(), raising=True)
    assert ocr_utils.extract_text_simple(str(image_path)) == "A"


@pytest.mark.unit
def test_extract_text_with_sizes_with_stub_ocr(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    image_path = tmp_path / "x.png"
    image_path.write_bytes(b"x")

    class StubOCR:
        def __call__(self, _path: str):
            return [
                [[[0, 0], [10, 0], [10, 16], [0, 16]], "A", 0.9],
                [[[0, 0], [10, 0], [10, 16], [0, 16]], "B", 0.4],
            ], [0.01]

    monkeypatch.setattr(ocr_utils, "get_ocr", lambda *_args, **_kwargs: StubOCR(), raising=True)
    blocks = ocr_utils.extract_text_with_sizes(str(image_path))
    assert blocks == [{"text": "A", "font_size": 12}]
