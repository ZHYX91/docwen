"""Focused tests split from test_ocr.py."""

from __future__ import annotations

from ._ocr_support import (
    Path,
    _write_ocr_input,
    pytest,
    sys,
    types,
)

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "status",
    [
        "no_text",
        "input_missing",
        "unavailable",
        "model_missing",
        "initialization_failed",
        "recognition_failed",
    ],
)
def test_ocr_outcome_exposes_text_only_for_success(status: str) -> None:
    import docwen_core.text.ocr as ocr

    assert ocr.OcrOutcome(ocr.OcrStatus(status), text="must stay hidden").recognized_text == ""
    assert ocr.OcrOutcome(ocr.OcrStatus.SUCCESS, text="recognized").recognized_text == "recognized"


def test_run_ocr_outcome_preserves_cached_unavailable_status(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", None)
    ocr.reset_ocr()
    image_path = _write_ocr_input(tmp_path)
    try:
        first = ocr.run_ocr_outcome(image_path, source_format="png", model_dir=tmp_path)
        second = ocr.run_ocr_outcome(image_path, source_format="png", model_dir=tmp_path)
    finally:
        ocr.reset_ocr()

    assert first.status is ocr.OcrStatus.UNAVAILABLE
    assert second.status is ocr.OcrStatus.UNAVAILABLE
    assert "not installed" in first.message
    assert second.message == first.message


def test_run_ocr_outcome_preserves_cached_initialization_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr
    from docwen_core.text import OCR_LANGUAGE_MODELS

    model_dir = tmp_path / "models"
    model_dir.mkdir()
    for name in OCR_LANGUAGE_MODELS["chinese"].values():
        (model_dir / name).write_bytes(b"onnx")
    initialization_count = 0

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            nonlocal initialization_count
            initialization_count += 1
            raise RuntimeError("initializer exploded")

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    ocr.reset_ocr()
    image_path = _write_ocr_input(tmp_path)
    try:
        first = ocr.run_ocr_outcome(
            image_path,
            source_format="png",
            ocr_language="chinese",
            model_dir=model_dir,
        )
        second = ocr.run_ocr_outcome(
            image_path,
            source_format="png",
            ocr_language="chinese",
            model_dir=model_dir,
        )
    finally:
        ocr.reset_ocr()

    assert initialization_count == 1
    assert first.status is ocr.OcrStatus.INITIALIZATION_FAILED
    assert second.status is ocr.OcrStatus.INITIALIZATION_FAILED
    assert first.message == "initializer exploded"
    assert second.message == first.message


def test_run_ocr_outcome_classifies_constructor_file_error_as_initialization_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr
    from docwen_core.text import OCR_LANGUAGE_MODELS

    model_dir = tmp_path / "models"
    model_dir.mkdir()
    for name in OCR_LANGUAGE_MODELS["chinese"].values():
        (model_dir / name).write_bytes(b"onnx")

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            raise FileNotFoundError("constructor resource missing")

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    ocr.reset_ocr()
    image_path = _write_ocr_input(tmp_path)
    try:
        outcome = ocr.run_ocr_outcome(
            image_path,
            source_format="png",
            ocr_language="chinese",
            model_dir=model_dir,
        )
    finally:
        ocr.reset_ocr()

    assert outcome.status is ocr.OcrStatus.INITIALIZATION_FAILED
    assert outcome.message == "constructor resource missing"


@pytest.mark.parametrize("error_type", [PermissionError, FileNotFoundError])
def test_run_ocr_outcome_classifies_model_directory_resolution_error_as_initialization_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    error_type: type[OSError],
) -> None:
    import docwen_core.text.ocr as ocr

    def _deny_model_directory(*_args: object, **_kwargs: object) -> Path:
        raise error_type("model directory access denied")

    monkeypatch.setattr(ocr, "_resolve_model_dir", _deny_model_directory)

    outcome = ocr.run_ocr_outcome(_write_ocr_input(tmp_path), source_format="png")

    assert outcome.status is ocr.OcrStatus.INITIALIZATION_FAILED
    assert outcome.text == ""
    assert outcome.message == "model directory access denied"


def test_run_ocr_outcome_classifies_malformed_result_as_recognition_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    class _Engine:
        def __call__(self, _path: str) -> tuple[int, float]:
            return 42, 0.1

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: ocr._OcrEngineSlot(_Engine()))

    outcome = ocr.run_ocr_outcome(_write_ocr_input(tmp_path), source_format="png")

    assert outcome.status is ocr.OcrStatus.RECOGNITION_FAILED
    assert outcome.text == ""
    assert "iterable" in outcome.message


def test_run_ocr_outcome_parses_box_text_score_result_shape(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    class _Engine:
        def __call__(self, _path: str) -> tuple[list[list[object]], list[float]]:
            return (
                [
                    [
                        [[72.0, 102.0], [1092.0, 102.0], [1092.0, 184.0], [72.0, 184.0]],
                        " HELLO DOCWEN OCR ",
                        0.98,
                    ]
                ],
                [0.1, 0.2, 0.3],
            )

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: ocr._OcrEngineSlot(_Engine()))

    outcome = ocr.run_ocr_outcome(_write_ocr_input(tmp_path), source_format="png")

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == "HELLO DOCWEN OCR"


def test_run_ocr_outcome_accepts_text_score_result_shape(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    class _Engine:
        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            return ([(" first line ", 0.99), ("second line", 0.98), (" ", 0.1)], 0.1)

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: ocr._OcrEngineSlot(_Engine()))

    outcome = ocr.run_ocr_outcome(_write_ocr_input(tmp_path), source_format="png")

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == "first line\nsecond line"


def test_run_ocr_outcome_filters_low_confidence_results(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    import docwen_core.text.ocr as ocr

    class _Engine:
        def __call__(self, _path: str) -> tuple[list[list[object] | tuple[str, object]], float]:
            return (
                [
                    [[[0, 0], [1, 0], [1, 1], [0, 1]], "box trusted", 0.500001],
                    [[[0, 0], [1, 0], [1, 1], [0, 1]], "box at threshold", 0.5],
                    ("compact trusted", "0.75"),
                    ("compact low", "0.49"),
                    ("compact nan", "nan"),
                    ("compact unknown", object()),
                ],
                0.1,
            )

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: ocr._OcrEngineSlot(_Engine()))

    outcome = ocr.run_ocr_outcome(_write_ocr_input(tmp_path), source_format="png")

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == "box trusted\ncompact trusted"
