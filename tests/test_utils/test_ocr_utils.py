from __future__ import annotations

import threading
from pathlib import Path

import pytest

import docwen.utils.ocr_utils as ocr_utils


@pytest.mark.unit
def test_resolve_ocr_language_auto_by_locale(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ocr_utils, "get_configured_ocr_language", lambda: ocr_utils.OCR_LANGUAGE_AUTO, raising=True)
    monkeypatch.setattr(ocr_utils, "LOCALE_TO_OCR_LANGUAGE", {"en_US": ocr_utils.OCR_LANGUAGE_ENGLISH}, raising=True)

    import docwen.i18n as i18n

    monkeypatch.setattr(i18n, "get_current_locale", lambda: "en_US", raising=True)
    assert ocr_utils.resolve_ocr_language() == ocr_utils.OCR_LANGUAGE_ENGLISH


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

