"""Focused tests split from test_ocr.py."""

from __future__ import annotations

from ._ocr_support import (
    Path,
    _first_existing_font,
    _write_ocr_input,
    pytest,
    sys,
    types,
)

pytestmark = pytest.mark.unit


def test_run_ocr_outcome_reports_missing_input_without_initializing_engine(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: pytest.fail("missing input must stop first"))

    outcome = ocr.run_ocr_outcome("__definitely_missing__.png", source_format="png")

    assert outcome.status is ocr.OcrStatus.INPUT_MISSING
    assert outcome.text == ""
    assert "__definitely_missing__.png" in outcome.message


def test_run_ocr_outcome_reports_actual_heic_unavailable_before_engine_init(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    image_path = tmp_path / "sample.heic"
    image_path.write_bytes(b"not a real heic")
    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: pytest.fail("HEIC should not initialize OCR"))

    outcome = ocr.run_ocr_outcome(image_path, source_format="heic")

    assert outcome.status is ocr.OcrStatus.UNAVAILABLE
    assert "pre-convert" in outcome.message


@pytest.mark.parametrize(
    ("suffix", "source_format", "payload"),
    [
        (".heic", "png", b"\x89PNG\r\n\x1a\ncontent-owned-format"),
        (".bin", "jpeg", b"\xff\xd8\xff\xe0content-owned-format"),
    ],
)
def test_run_ocr_outcome_uses_content_format_not_suffix(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    suffix: str,
    source_format: str,
    payload: bytes,
) -> None:
    import docwen_core.text.ocr as ocr

    image_path = tmp_path / f"misnamed{suffix}"
    image_path.write_bytes(payload)
    calls: list[str] = []

    class _Engine:
        def __call__(self, path: str) -> tuple[list[tuple[str, float]], float]:
            calls.append(path)
            return ([("content format won", 0.99)], 0.01)

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: ocr._OcrEngineSlot(_Engine()))

    outcome = ocr.run_ocr_outcome(image_path, source_format=source_format)

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == "content format won"
    assert calls == [str(image_path)]


@pytest.mark.parametrize("source_format", ["", "image", "pdf", "unknown"])
def test_run_ocr_rejects_non_concrete_image_formats(source_format: str) -> None:
    import docwen_core.text.ocr as ocr

    with pytest.raises(ValueError, match="concrete image format"):
        ocr.run_ocr_outcome("sample.png", source_format=source_format)


def test_run_ocr_initializes_rapidocr_with_selected_language_models(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "models" / "rapidocr"
    model_dir.mkdir(parents=True)
    for name in (
        "ch_PP-OCRv4_det_infer.onnx",
        "japan_PP-OCRv4_rec_infer.onnx",
        "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    ):
        (model_dir / name).write_bytes(b"onnx")

    init_kwargs: list[dict[str, str]] = []
    onnxruntime_severities: list[int] = []

    class _RapidOCR:
        def __init__(self, **kwargs: str) -> None:
            init_kwargs.append(kwargs)

        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            return ([("日本語", 0.99)], 0.1)

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    monkeypatch.setitem(
        sys.modules,
        "onnxruntime",
        types.SimpleNamespace(set_default_logger_severity=onnxruntime_severities.append),
    )
    monkeypatch.setenv("DOCWEN_RAPIDOCR_MODEL_DIR", str(model_dir))
    ocr.reset_ocr()
    image_path = _write_ocr_input(tmp_path)

    outcome = ocr.run_ocr_outcome(image_path, source_format="png", ocr_language="japanese")

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == "日本語"
    assert onnxruntime_severities == [3]
    assert len(init_kwargs) == 1
    assert init_kwargs[0]["det_model_path"] == str(model_dir / "ch_PP-OCRv4_det_infer.onnx")
    assert init_kwargs[0]["rec_model_path"] == str(model_dir / "japan_PP-OCRv4_rec_infer.onnx")
    assert init_kwargs[0]["cls_model_path"] == str(model_dir / "ch_ppocr_mobile_v2.0_cls_infer.onnx")


@pytest.mark.parametrize(
    ("ocr_language", "expected_rec_model"),
    [
        ("chinese", "ch_PP-OCRv4_rec_infer.onnx"),
        ("chinese_cht", "chinese_cht_PP-OCRv3_rec_infer.onnx"),
        ("english", "en_PP-OCRv4_rec_infer.onnx"),
        ("japanese", "japan_PP-OCRv4_rec_infer.onnx"),
        ("korean", "korean_PP-OCRv4_rec_infer.onnx"),
        ("latin", "latin_PP-OCRv3_rec_infer.onnx"),
        ("cyrillic", "cyrillic_PP-OCRv3_rec_infer.onnx"),
    ],
)
def test_run_ocr_initializes_rapidocr_for_each_configured_language_model(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    ocr_language: str,
    expected_rec_model: str,
) -> None:
    import docwen_core.text.ocr as ocr
    from docwen_core.text import OCR_LANGUAGE_MODELS

    model_dir = tmp_path / "models" / "rapidocr"
    model_dir.mkdir(parents=True)
    for models in OCR_LANGUAGE_MODELS.values():
        for name in models.values():
            (model_dir / name).write_bytes(b"onnx")

    init_kwargs: list[dict[str, str]] = []

    class _RapidOCR:
        def __init__(self, **kwargs: str) -> None:
            init_kwargs.append(kwargs)

        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            return ([(ocr_language, 0.99)], 0.1)

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    monkeypatch.setenv("DOCWEN_RAPIDOCR_MODEL_DIR", str(model_dir))
    ocr.reset_ocr()
    image_path = _write_ocr_input(tmp_path)

    outcome = ocr.run_ocr_outcome(image_path, source_format="png", ocr_language=ocr_language)

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == ocr_language
    assert len(init_kwargs) == 1
    assert init_kwargs[0]["det_model_path"] == str(model_dir / "ch_PP-OCRv4_det_infer.onnx")
    assert init_kwargs[0]["rec_model_path"] == str(model_dir / expected_rec_model)
    assert init_kwargs[0]["cls_model_path"] == str(model_dir / "ch_ppocr_mobile_v2.0_cls_infer.onnx")


@pytest.mark.parametrize(
    ("locale", "expected_text", "expected_rec_model"),
    [
        ("zh_CN", "chinese", "ch_PP-OCRv4_rec_infer.onnx"),
        ("zh_TW", "chinese_cht", "chinese_cht_PP-OCRv3_rec_infer.onnx"),
        ("en_US", "english", "en_PP-OCRv4_rec_infer.onnx"),
        ("ja_JP", "japanese", "japan_PP-OCRv4_rec_infer.onnx"),
        ("ko_KR", "korean", "korean_PP-OCRv4_rec_infer.onnx"),
        ("de_DE", "latin", "latin_PP-OCRv3_rec_infer.onnx"),
        ("fr_FR", "latin", "latin_PP-OCRv3_rec_infer.onnx"),
        ("pt_BR", "latin", "latin_PP-OCRv3_rec_infer.onnx"),
        ("es_ES", "latin", "latin_PP-OCRv3_rec_infer.onnx"),
        ("vi_VN", "latin", "latin_PP-OCRv3_rec_infer.onnx"),
        ("ru_RU", "cyrillic", "cyrillic_PP-OCRv3_rec_infer.onnx"),
    ],
)
def test_run_ocr_auto_locale_selects_expected_language_model(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    locale: str,
    expected_text: str,
    expected_rec_model: str,
) -> None:
    import docwen_core.text.ocr as ocr
    from docwen_core.text import OCR_LANGUAGE_MODELS

    model_dir = tmp_path / "models" / "rapidocr"
    model_dir.mkdir(parents=True)
    for models in OCR_LANGUAGE_MODELS.values():
        for name in models.values():
            (model_dir / name).write_bytes(b"onnx")

    init_kwargs: list[dict[str, str]] = []

    class _RapidOCR:
        def __init__(self, **kwargs: str) -> None:
            init_kwargs.append(kwargs)

        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            return ([(expected_text, 0.99)], 0.1)

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    monkeypatch.setenv("DOCWEN_RAPIDOCR_MODEL_DIR", str(model_dir))
    ocr.reset_ocr()
    image_path = _write_ocr_input(tmp_path)

    outcome = ocr.run_ocr_outcome(
        image_path,
        source_format="png",
        ocr_language="auto",
        current_locale=locale,
    )

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == expected_text
    assert len(init_kwargs) == 1
    assert init_kwargs[0]["rec_model_path"] == str(model_dir / expected_rec_model)


def test_run_ocr_real_japanese_smoke_uses_selected_language_model(tmp_path: Path) -> None:
    pytest.importorskip("rapidocr_onnxruntime")
    from PIL import Image, ImageDraw, ImageFont

    import docwen_core.text.ocr as ocr

    font_path = _first_existing_font(
        (
            r"C:\Windows\Fonts\YuGothB.ttc",
            r"C:\Windows\Fonts\msgothic.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc",
        )
    )
    if font_path is None:
        pytest.skip("No Japanese-capable font available for real OCR smoke")

    image = Image.new("RGB", (1000, 260), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(str(font_path), 82)
    draw.text((40, 80), "日本語 OCR", fill="black", font=font)
    image_path = tmp_path / "japanese_ocr.png"
    image.save(image_path)

    ocr.reset_ocr()
    try:
        outcome = ocr.run_ocr_outcome(
            image_path,
            source_format="png",
            ocr_language="japanese",
            current_locale="ja_JP",
        )
    finally:
        ocr.reset_ocr()

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == "日本語 OCR"


def test_run_ocr_real_english_degraded_scan_smoke(tmp_path: Path) -> None:
    pytest.importorskip("rapidocr_onnxruntime")
    from PIL import Image, ImageDraw, ImageFilter, ImageFont

    import docwen_core.text.ocr as ocr

    font_path = _first_existing_font(
        (
            r"C:\Windows\Fonts\arialbd.ttf",
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\segoeuib.ttf",
            r"C:\Windows\Fonts\segoeui.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/Library/Fonts/Arial Bold.ttf",
            "/Library/Fonts/Arial.ttf",
        )
    )
    if font_path is None:
        pytest.skip("No Latin font available for degraded OCR smoke")

    image = Image.new("RGB", (1400, 360), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(str(font_path), 96)
    draw.text((70, 120), "DOCWEN OCR 2026", fill=(10, 10, 10), font=font)
    for y in range(0, image.height, 8):
        draw.line((0, y, image.width, y), fill=(230, 230, 230), width=1)
    image = image.filter(ImageFilter.GaussianBlur(0.6)).rotate(1.0, expand=True, fillcolor="white")
    image_path = tmp_path / "english_degraded_scan_ocr.png"
    image.save(image_path)

    ocr.reset_ocr()
    try:
        outcome = ocr.run_ocr_outcome(
            image_path,
            source_format="png",
            ocr_language="english",
            current_locale="en_US",
        )
    finally:
        ocr.reset_ocr()

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == "DOCWEN OCR 2026"


def test_run_ocr_real_additional_language_token_smoke(tmp_path: Path) -> None:
    pytest.importorskip("rapidocr_onnxruntime")
    from PIL import Image, ImageDraw, ImageFont

    import docwen_core.text.ocr as ocr

    cases = (
        (
            "korean",
            "ko_KR",
            "문서",
            "문서",
            (
                r"C:\Windows\Fonts\malgunbd.ttf",
                r"C:\Windows\Fonts\malgun.ttf",
                "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
                "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
                "/System/Library/Fonts/AppleSDGothicNeo.ttc",
            ),
        ),
        (
            "cyrillic",
            "ru_RU",
            "ДОК",
            "док",
            (
                r"C:\Windows\Fonts\arialbd.ttf",
                r"C:\Windows\Fonts\arial.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "/Library/Fonts/Arial Bold.ttf",
                "/Library/Fonts/Arial.ttf",
            ),
        ),
        (
            "latin",
            "de_DE",
            "DOCWEN",
            "docwen",
            (
                r"C:\Windows\Fonts\arialbd.ttf",
                r"C:\Windows\Fonts\arial.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "/Library/Fonts/Arial Bold.ttf",
                "/Library/Fonts/Arial.ttf",
            ),
        ),
        (
            "chinese_cht",
            "zh_TW",
            "繁體",
            "繁體",
            (
                r"C:\Windows\Fonts\msjhbd.ttc",
                r"C:\Windows\Fonts\msjh.ttc",
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                "/System/Library/Fonts/PingFang.ttc",
            ),
        ),
    )
    missing_fonts = [
        language for language, _locale, _text, _expected, fonts in cases if _first_existing_font(fonts) is None
    ]
    if missing_fonts:
        pytest.skip(f"Missing fonts for additional OCR language smoke: {', '.join(missing_fonts)}")

    for language, locale, source_text, expected_text, fonts in cases:
        font_path = _first_existing_font(fonts)
        assert font_path is not None
        image = Image.new("RGB", (1000, 320), "white")
        draw = ImageDraw.Draw(image)
        font = ImageFont.truetype(str(font_path), 96)
        bbox = draw.textbbox((0, 0), source_text, font=font)
        x = max(30, (image.width - (bbox[2] - bbox[0])) // 2)
        y = max(30, (image.height - (bbox[3] - bbox[1])) // 2 - 10)
        draw.text((x, y), source_text, fill="black", font=font)
        image_path = tmp_path / f"{language}_token_ocr.png"
        image.save(image_path)

        ocr.reset_ocr()
        try:
            outcome = ocr.run_ocr_outcome(
                image_path,
                source_format="png",
                ocr_language=language,
                current_locale=locale,
            )
        finally:
            ocr.reset_ocr()

        assert outcome.status is ocr.OcrStatus.SUCCESS
        assert outcome.text.casefold() == expected_text


def test_run_ocr_outcome_reports_missing_model(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    import docwen_core.text.ocr as ocr

    model_dir = tmp_path / "models" / "rapidocr"
    model_dir.mkdir(parents=True)
    for name in (
        "ch_PP-OCRv4_det_infer.onnx",
        "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    ):
        (model_dir / name).write_bytes(b"onnx")

    class _RapidOCR:
        def __init__(self, **_kwargs: str) -> None:
            raise AssertionError("RapidOCR should not initialize when a required model is missing")

    monkeypatch.setitem(sys.modules, "rapidocr_onnxruntime", types.SimpleNamespace(RapidOCR=_RapidOCR))
    monkeypatch.setenv("DOCWEN_RAPIDOCR_MODEL_DIR", str(model_dir))
    ocr.reset_ocr()
    image_path = _write_ocr_input(tmp_path)

    outcome = ocr.run_ocr_outcome(image_path, source_format="png", ocr_language="japanese")

    assert outcome.status is ocr.OcrStatus.MODEL_MISSING
    assert outcome.text == ""
    assert "japanese" in outcome.message
    assert "OCR models missing" in caplog.text


def test_run_ocr_outcome_distinguishes_recognition_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    class _Engine:
        def __call__(self, _path: str) -> tuple[None, float]:
            raise RuntimeError("detector exploded")

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: ocr._OcrEngineSlot(_Engine()))

    outcome = ocr.run_ocr_outcome(_write_ocr_input(tmp_path), source_format="png")

    assert outcome.status is ocr.OcrStatus.RECOGNITION_FAILED
    assert outcome.text == ""
    assert "detector exploded" in outcome.message


def test_run_ocr_outcome_distinguishes_no_text(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import docwen_core.text.ocr as ocr

    class _Engine:
        def __call__(self, _path: str) -> tuple[None, float]:
            return None, 0.1

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: ocr._OcrEngineSlot(_Engine()))

    outcome = ocr.run_ocr_outcome(_write_ocr_input(tmp_path, "blank.png"), source_format="png")

    assert outcome.status is ocr.OcrStatus.NO_TEXT
    assert outcome.text == ""
    assert outcome.message == ""


def test_run_ocr_outcome_reports_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import docwen_core.text.ocr as ocr

    class _Engine:
        def __call__(self, _path: str) -> tuple[list[tuple[str, float]], float]:
            return ([(" recognized ", 0.99)], 0.1)

    monkeypatch.setattr(ocr, "_get_ocr_slot", lambda *_args, **_kwargs: ocr._OcrEngineSlot(_Engine()))

    outcome = ocr.run_ocr_outcome(_write_ocr_input(tmp_path), source_format="png")

    assert outcome.status is ocr.OcrStatus.SUCCESS
    assert outcome.text == "recognized"
    assert outcome.message == ""


@pytest.mark.parametrize(
    ("status", "expected"),
    [
        (
            "success",
            "OCR best-effort result: status=success; OCR text is machine-generated and may contain "
            "recognition errors or omissions; verify it against the source.",
        ),
        (
            "no_text",
            "OCR best-effort result: status=no_text; OCR detected no text; text may have been missed, "
            "so verify it against the source.",
        ),
        (
            "input_missing",
            "OCR best-effort fallback: status=input_missing; OCR input file is missing or is not a regular file.",
        ),
        ("unavailable", "OCR best-effort fallback: status=unavailable; OCR engine is unavailable."),
        ("model_missing", "OCR best-effort fallback: status=model_missing; OCR model files are missing."),
        (
            "initialization_failed",
            "OCR best-effort fallback: status=initialization_failed; OCR engine initialization failed.",
        ),
        ("recognition_failed", "OCR best-effort fallback: status=recognition_failed; OCR recognition failed."),
    ],
)
def test_format_ocr_best_effort_warning_is_canonical_and_safe(status: str, expected: str) -> None:
    import docwen_core.text.ocr as ocr

    assert ocr.format_ocr_best_effort_warning(ocr.OcrStatus(status)) == expected
    assert ocr.format_ocr_best_effort_warning(status) == expected
    assert ocr.format_ocr_best_effort_warning(status, context="html image sample.png") == (
        f"{expected[:-1]}; html image sample.png."
    )


@pytest.mark.parametrize("status", ["unknown", object()])
def test_format_ocr_best_effort_warning_ignores_unknown_status(status: object) -> None:
    import docwen_core.text.ocr as ocr

    assert ocr.format_ocr_best_effort_warning(status) is None
