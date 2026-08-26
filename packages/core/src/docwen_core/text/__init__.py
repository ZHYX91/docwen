"""OCR language resolution helpers.

These are pure-data / pure-function utilities with **zero** dependencies
on RapidOCR, runtime config, or GUI i18n.  Plugins feed explicit
``ocr_language`` / ``current_locale`` values so that the resolution
logic stays testable without any I/O.
"""

from __future__ import annotations

# ═══════════════════════════════════════════════════════════════════════════
# OCR language constants
# ═══════════════════════════════════════════════════════════════════════════

OCR_LANGUAGE_AUTO: str = "auto"
OCR_LANGUAGE_CHINESE: str = "chinese"
OCR_LANGUAGE_CHINESE_CHT: str = "chinese_cht"
OCR_LANGUAGE_ENGLISH: str = "english"
OCR_LANGUAGE_JAPANESE: str = "japanese"
OCR_LANGUAGE_KOREAN: str = "korean"
OCR_LANGUAGE_LATIN: str = "latin"
OCR_LANGUAGE_CYRILLIC: str = "cyrillic"

# ═══════════════════════════════════════════════════════════════════════════
# Locale → OCR language mapping
# ═══════════════════════════════════════════════════════════════════════════

LOCALE_TO_OCR_LANGUAGE: dict[str, str] = {
    "zh_CN": OCR_LANGUAGE_CHINESE,
    "zh_TW": OCR_LANGUAGE_CHINESE_CHT,
    "en_US": OCR_LANGUAGE_ENGLISH,
    "ja_JP": OCR_LANGUAGE_JAPANESE,
    "ko_KR": OCR_LANGUAGE_KOREAN,
    "de_DE": OCR_LANGUAGE_LATIN,
    "fr_FR": OCR_LANGUAGE_LATIN,
    "pt_BR": OCR_LANGUAGE_LATIN,
    "es_ES": OCR_LANGUAGE_LATIN,
    "vi_VN": OCR_LANGUAGE_LATIN,
    "ru_RU": OCR_LANGUAGE_CYRILLIC,
}

# ═══════════════════════════════════════════════════════════════════════════
# OCR language → ONNX model files (RapidOCR)
# ═══════════════════════════════════════════════════════════════════════════

OCR_LANGUAGE_MODELS: dict[str, dict[str, str]] = {
    OCR_LANGUAGE_CHINESE: {
        "det": "ch_PP-OCRv4_det_infer.onnx",
        "rec": "ch_PP-OCRv4_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    },
    OCR_LANGUAGE_CHINESE_CHT: {
        "det": "ch_PP-OCRv4_det_infer.onnx",
        "rec": "chinese_cht_PP-OCRv3_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    },
    OCR_LANGUAGE_ENGLISH: {
        "det": "ch_PP-OCRv4_det_infer.onnx",
        "rec": "en_PP-OCRv4_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    },
    OCR_LANGUAGE_JAPANESE: {
        "det": "ch_PP-OCRv4_det_infer.onnx",
        "rec": "japan_PP-OCRv4_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    },
    OCR_LANGUAGE_KOREAN: {
        "det": "ch_PP-OCRv4_det_infer.onnx",
        "rec": "korean_PP-OCRv4_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    },
    OCR_LANGUAGE_LATIN: {
        "det": "ch_PP-OCRv4_det_infer.onnx",
        "rec": "latin_PP-OCRv3_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    },
    OCR_LANGUAGE_CYRILLIC: {
        "det": "ch_PP-OCRv4_det_infer.onnx",
        "rec": "cyrillic_PP-OCRv3_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    },
}


# ═══════════════════════════════════════════════════════════════════════════
# Public helpers
# ═══════════════════════════════════════════════════════════════════════════


def resolve_ocr_language(
    ocr_language: str,
    current_locale: str = "zh_CN",
) -> str:
    """Resolve the effective OCR language from config + locale.

    When *ocr_language* is ``"auto"`` the function maps *current_locale*
    through :data:`LOCALE_TO_OCR_LANGUAGE`; otherwise the explicit value
    is returned unchanged.

    Args:
        ocr_language: Configured OCR language (``"auto"``, ``"chinese"``,
            ``"japanese"``, …).
        current_locale: Active UI locale (e.g. ``"zh_CN"``, ``"ja_JP"``).
            Only used when *ocr_language* is ``"auto"``.

    Returns:
        A concrete OCR language key present in :data:`OCR_LANGUAGE_MODELS`.
    """
    if ocr_language == OCR_LANGUAGE_AUTO:
        return LOCALE_TO_OCR_LANGUAGE.get(current_locale, OCR_LANGUAGE_CHINESE)
    return ocr_language
