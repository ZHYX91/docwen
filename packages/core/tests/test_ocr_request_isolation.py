"""Request-aware availability and shared-engine isolation tests for OCR."""

from __future__ import annotations

import sys
import threading
import time
import types
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit


def _write_language_models(model_dir: Path, *languages: str) -> None:
    from docwen_core.text import OCR_LANGUAGE_MODELS

    model_dir.mkdir(parents=True, exist_ok=True)
    for language in languages:
        for filename in OCR_LANGUAGE_MODELS[language].values():
            (model_dir / filename).write_bytes(b"onnx")


def _write_ocr_input(tmp_path: Path, name: str) -> Path:
    image_path = tmp_path / name
    image_path.write_bytes(b"ocr input")
    return image_path


def test_ocr_available_uses_requested_language_and_locale(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "japanese-only"
    _write_language_models(model_dir, "japanese")
    recognition_models: list[str] = []

    class _RapidOCR:
        def __init__(self, **kwargs: str) -> None:
            recognition_models.append(Path(kwargs["rec_model_path"]).name)

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    ocr.reset_ocr()
    try:
        assert not ocr.ocr_available(model_dir=model_dir)
        assert ocr.ocr_available(
            ocr_language="auto",
            current_locale="ja_JP",
            model_dir=model_dir,
        )
    finally:
        ocr.reset_ocr()

    assert recognition_models == ["japan_PP-OCRv4_rec_infer.onnx"]


def test_ocr_available_rejects_missing_requested_models_even_when_default_exists(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "chinese-only"
    _write_language_models(model_dir, "chinese")

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            pass

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    ocr.reset_ocr()
    try:
        assert ocr.ocr_available(
            ocr_language="chinese",
            current_locale="zh_CN",
            model_dir=model_dir,
        )
        assert not ocr.ocr_available(
            ocr_language="japanese",
            current_locale="ja_JP",
            model_dir=model_dir,
        )
    finally:
        ocr.reset_ocr()


def test_run_ocr_serializes_invocations_of_the_same_cached_engine(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "models"
    _write_language_models(model_dir, "chinese")
    first_path = _write_ocr_input(tmp_path, "first.png")
    second_path = _write_ocr_input(tmp_path, "second.png")
    state_lock = threading.Lock()
    first_call_entered = threading.Event()
    allow_first_call_to_finish = threading.Event()
    second_slot_obtained = threading.Event()
    second_call_entered = threading.Event()
    active_calls = 0
    call_count = 0
    initialization_count = 0
    lookup_count = 0
    max_active_calls = 0

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            nonlocal initialization_count
            initialization_count += 1

        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            nonlocal active_calls, call_count, max_active_calls
            with state_lock:
                call_count += 1
                call_number = call_count
                active_calls += 1
                max_active_calls = max(max_active_calls, active_calls)
                if call_number == 1:
                    first_call_entered.set()
                else:
                    second_call_entered.set()

            if call_number == 1:
                assert allow_first_call_to_finish.wait(timeout=5.0)

            with state_lock:
                active_calls -= 1
            return ([("OCR", 0.99)], 0.01)

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    original_get_ocr_slot = ocr._get_ocr_slot

    def _tracked_get_ocr_slot(*args: Any, **kwargs: Any) -> Any:
        nonlocal lookup_count
        slot = original_get_ocr_slot(*args, **kwargs)
        with state_lock:
            lookup_count += 1
            if lookup_count == 2:
                second_slot_obtained.set()
        return slot

    monkeypatch.setattr(ocr, "_get_ocr_slot", _tracked_get_ocr_slot)
    ocr.reset_ocr()
    try:
        with ThreadPoolExecutor(max_workers=2) as executor:
            first = executor.submit(
                ocr.run_ocr_outcome,
                first_path,
                "chinese",
                source_format="png",
                model_dir=model_dir,
            )
            assert first_call_entered.wait(timeout=1.0)
            second = executor.submit(
                ocr.run_ocr_outcome,
                second_path,
                "chinese",
                source_format="png",
                model_dir=model_dir,
            )
            assert second_slot_obtained.wait(timeout=2.0)
            second_entered_before_release = second_call_entered.wait(timeout=0.5)
            allow_first_call_to_finish.set()
            assert first.result(timeout=5.0).recognized_text == "OCR"
            assert second.result(timeout=5.0).recognized_text == "OCR"
    finally:
        allow_first_call_to_finish.set()
        ocr.reset_ocr()

    assert call_count == 2
    assert initialization_count == 1
    assert lookup_count == 2
    assert not second_entered_before_release
    assert max_active_calls == 1


def test_run_ocr_allows_different_cached_engines_to_run_in_parallel(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "models"
    _write_language_models(model_dir, "chinese", "japanese")
    chinese_path = _write_ocr_input(tmp_path, "zh.png")
    japanese_path = _write_ocr_input(tmp_path, "ja.png")
    state_lock = threading.Lock()
    invocation_barrier = threading.Barrier(2)
    active_calls = 0
    max_active_calls = 0

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            pass

        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            nonlocal active_calls, max_active_calls
            with state_lock:
                active_calls += 1
                max_active_calls = max(max_active_calls, active_calls)

            invocation_barrier.wait(timeout=2.0)

            with state_lock:
                active_calls -= 1
            return ([("OCR", 0.99)], 0.01)

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    ocr.reset_ocr()
    try:
        with ThreadPoolExecutor(max_workers=2) as executor:
            chinese = executor.submit(
                ocr.run_ocr_outcome,
                chinese_path,
                "chinese",
                source_format="png",
                model_dir=model_dir,
            )
            japanese = executor.submit(
                ocr.run_ocr_outcome,
                japanese_path,
                "japanese",
                source_format="png",
                model_dir=model_dir,
            )
            assert chinese.result(timeout=5.0).recognized_text == "OCR"
            assert japanese.result(timeout=5.0).recognized_text == "OCR"
    finally:
        ocr.reset_ocr()

    assert max_active_calls == 2


def test_reset_ocr_waits_for_in_flight_initialization_before_clearing_cache(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "models"
    _write_language_models(model_dir, "chinese")
    first_path = _write_ocr_input(tmp_path, "first.png")
    second_path = _write_ocr_input(tmp_path, "second.png")
    constructor_entered = threading.Event()
    allow_constructor_to_finish = threading.Event()
    reset_started = threading.Event()
    reset_finished = threading.Event()
    initialization_count = 0

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            nonlocal initialization_count
            initialization_count += 1
            if initialization_count == 1:
                constructor_entered.set()
                assert allow_constructor_to_finish.wait(timeout=2.0)

        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            return ([("OCR", 0.99)], 0.01)

    def _reset() -> None:
        reset_started.set()
        ocr.reset_ocr()
        reset_finished.set()

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    ocr.reset_ocr()
    try:
        with ThreadPoolExecutor(max_workers=2) as executor:
            initialization = executor.submit(
                ocr.run_ocr_outcome,
                first_path,
                "chinese",
                source_format="png",
                model_dir=model_dir,
            )
            assert constructor_entered.wait(timeout=1.0)
            reset = executor.submit(_reset)
            assert reset_started.wait(timeout=1.0)
            reset_completed_during_initialization = reset_finished.wait(timeout=1.0)
            allow_constructor_to_finish.set()
            assert initialization.result(timeout=5.0).recognized_text == "OCR"
            reset.result(timeout=5.0)

        assert (
            ocr.run_ocr_outcome(second_path, "chinese", source_format="png", model_dir=model_dir).recognized_text
            == "OCR"
        )
    finally:
        allow_constructor_to_finish.set()
        ocr.reset_ocr()

    assert not reset_completed_during_initialization
    assert initialization_count == 2


def test_reset_ocr_waits_for_active_invocation_before_clearing_cache(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "models"
    _write_language_models(model_dir, "chinese")
    first_path = _write_ocr_input(tmp_path, "first.png")
    second_path = _write_ocr_input(tmp_path, "second.png")
    invocation_entered = threading.Event()
    allow_invocation_to_finish = threading.Event()
    reset_started = threading.Event()
    reset_finished = threading.Event()
    initialization_count = 0

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            nonlocal initialization_count
            initialization_count += 1

        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            if initialization_count == 1:
                invocation_entered.set()
                assert allow_invocation_to_finish.wait(timeout=5.0)
            return ([("OCR", 0.99)], 0.01)

    def _reset() -> None:
        reset_started.set()
        ocr.reset_ocr()
        reset_finished.set()

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    ocr.reset_ocr()
    try:
        with ThreadPoolExecutor(max_workers=2) as executor:
            invocation = executor.submit(
                ocr.run_ocr_outcome,
                first_path,
                "chinese",
                source_format="png",
                model_dir=model_dir,
            )
            assert invocation_entered.wait(timeout=1.0)
            reset = executor.submit(_reset)
            assert reset_started.wait(timeout=1.0)
            reset_completed_during_invocation = reset_finished.wait(timeout=0.5)
            allow_invocation_to_finish.set()
            assert invocation.result(timeout=5.0).recognized_text == "OCR"
            reset.result(timeout=5.0)

        assert (
            ocr.run_ocr_outcome(second_path, "chinese", source_format="png", model_dir=model_dir).recognized_text
            == "OCR"
        )
    finally:
        allow_invocation_to_finish.set()
        ocr.reset_ocr()

    assert not reset_completed_during_invocation
    assert initialization_count == 2


def test_reset_ocr_drains_all_active_slots_and_blocks_new_lookup(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "models"
    _write_language_models(model_dir, "chinese", "japanese")
    chinese_path = _write_ocr_input(tmp_path, "zh.png")
    japanese_path = _write_ocr_input(tmp_path, "ja.png")
    later_path = _write_ocr_input(tmp_path, "later.png")
    both_invocations_entered = threading.Barrier(3)
    allow_invocations_to_finish = threading.Event()
    reset_started = threading.Event()
    reset_finished = threading.Event()
    later_invocation_entered = threading.Event()
    state_lock = threading.Lock()
    initialization_count = 0

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            nonlocal initialization_count
            with state_lock:
                initialization_count += 1
                instance_number = initialization_count
            self._instance_number = instance_number

        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            if self._instance_number <= 2:
                both_invocations_entered.wait(timeout=2.0)
                assert allow_invocations_to_finish.wait(timeout=5.0)
            else:
                later_invocation_entered.set()
            return ([("OCR", 0.99)], 0.01)

    def _reset() -> None:
        reset_started.set()
        ocr.reset_ocr()
        reset_finished.set()

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    ocr.reset_ocr()
    try:
        with ThreadPoolExecutor(max_workers=4) as executor:
            chinese = executor.submit(
                ocr.run_ocr_outcome,
                chinese_path,
                "chinese",
                source_format="png",
                model_dir=model_dir,
            )
            japanese = executor.submit(
                ocr.run_ocr_outcome,
                japanese_path,
                "japanese",
                source_format="png",
                model_dir=model_dir,
            )
            both_invocations_entered.wait(timeout=2.0)

            reset = executor.submit(_reset)
            assert reset_started.wait(timeout=1.0)
            deadline = time.monotonic() + 1.0
            while not ocr._ocr_lock.locked() and time.monotonic() < deadline:
                time.sleep(0.005)
            assert ocr._ocr_lock.locked()
            later = executor.submit(
                ocr.run_ocr_outcome,
                later_path,
                "chinese",
                source_format="png",
                model_dir=model_dir,
            )
            lookup_completed_during_reset = later_invocation_entered.wait(timeout=0.5)
            reset_completed_during_invocations = reset_finished.is_set()

            allow_invocations_to_finish.set()
            assert chinese.result(timeout=5.0).recognized_text == "OCR"
            assert japanese.result(timeout=5.0).recognized_text == "OCR"
            reset.result(timeout=5.0)
            assert later.result(timeout=5.0).recognized_text == "OCR"
    finally:
        allow_invocations_to_finish.set()
        ocr.reset_ocr()

    assert not lookup_completed_during_reset
    assert not reset_completed_during_invocations
    assert initialization_count == 3
