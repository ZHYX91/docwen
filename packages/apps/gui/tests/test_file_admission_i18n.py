"""Locale contracts for structured file-admission diagnostics."""

from __future__ import annotations

import string
import tomllib
from pathlib import Path

import pytest

from docwen_core.models import FILE_INSPECTION_METADATA_KEY, FileInspection
from docwen_core.models.file_ref import FileRef
from docwen_gui.file_admission_i18n import (
    file_admission_message_key,
    render_file_admission_code,
    render_file_admission_message,
    render_file_format_notice,
    render_file_inspection_message,
    render_remaining_file_warnings,
)
from docwen_gui.i18n import get_locale, set_locale

pytestmark = pytest.mark.unit


def test_compact_format_notice_preserves_other_warning_and_reason() -> None:
    ref = FileRef(
        path="sample.docx",
        format="doc",
        category="document",
        warning_message="Another consumer warning",
        metadata={
            FILE_INSPECTION_METADATA_KEY: {
                "warning_code": "FILE_FORMAT_SAME_FAMILY_MISMATCH",
                "warning_message": "Format mismatch",
                "declared_format": "docx",
                "detected_format": "doc",
                "warnings": [
                    {"code": "FILE_FORMAT_SAME_FAMILY_MISMATCH", "message": "Format mismatch"},
                    {"code": "SIGNATURE_UNKNOWN", "message": "Signature unavailable"},
                ],
                "reason_code": "CONTENT_LIMITED",
                "reason_message": "Limited preview",
            }
        },
    )
    warning = render_remaining_file_warnings(ref)
    assert "Format mismatch" not in warning
    assert warning.splitlines() == ["Another consumer warning", "Signature unavailable", "Limited preview"]


_LOCALES_DIR = Path(__file__).resolve().parents[4] / "i18n" / "locales"
_LOCALES = (
    "de_DE",
    "en_US",
    "es_ES",
    "fr_FR",
    "ja_JP",
    "ko_KR",
    "pt_BR",
    "ru_RU",
    "vi_VN",
    "zh_CN",
    "zh_TW",
)


@pytest.mark.parametrize("locale", _LOCALES)
def test_compact_notice_filters_format_from_combined_localized_diagnostics(locale: str) -> None:
    original_locale = get_locale()
    try:
        set_locale(locale)
        inspection = _inspection(
            declared_format="docx",
            detected_format="doc",
            warning_code="FILE_FORMAT_SAME_FAMILY_MISMATCH",
            warning_message="Format mismatch. Signature unavailable.",
            warnings=(
                {"code": "FILE_FORMAT_SAME_FAMILY_MISMATCH", "message": "Format mismatch."},
                {"code": "SIGNATURE_UNKNOWN", "message": "Signature unavailable."},
            ),
        )
        ref = FileRef(
            path="sample.docx",
            format="doc",
            category="document",
            warning_message=render_file_inspection_message(inspection),
            metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
        )
        assert render_remaining_file_warnings(ref) == "[SIGNATURE_UNKNOWN] Signature unavailable."
    finally:
        set_locale(original_locale)


_MAIN_WINDOW_KEYS = frozenset(
    {
        "file_admission_invalid",
        "file_admission_missing",
        "file_admission_unreadable",
        "file_admission_changed",
        "file_admission_blocked",
        "file_admission_confirm_title",
        "file_admission_confirm_message",
        "file_admission_confirm_action",
        "file_admission_cancelled",
    }
)
_CODE_TO_KEY = {
    "FILE_FORMAT_COMPATIBLE_TEXT": "compatible_text",
    "FILE_FORMAT_SAME_FAMILY_MISMATCH": "same_family_mismatch",
    "FILE_FORMAT_CROSS_FAMILY_MISMATCH": "cross_family_mismatch",
    "FILE_EXTENSION_UNSUPPORTED": "unknown_extension",
    "FILE_EMPTY": "empty",
    "FILE_CONTAINER_INVALID": "container_invalid",
    "FILE_CONTAINER_UNSUPPORTED": "container_unsupported",
    "FILE_CONTAINER_UNRECOGNIZED": "container_unrecognized",
    "FILE_CONTENT_UNRECOGNIZED": "content_unrecognized",
    "FILE_READ_ERROR": "read_error",
    "UNSUPPORTED_FORMAT": "unsupported_format",
}


def _load(locale: str) -> dict[str, object]:
    return tomllib.loads((_LOCALES_DIR / f"{locale}.toml").read_text(encoding="utf-8"))


def _inspection(**overrides: object) -> FileInspection:
    payload: dict[str, object] = {
        "declared_format": "docx",
        "detected_format": "pdf",
        "warning_code": "",
        "warning_message": "",
        "reason_code": "",
        "reason_message": "",
        "warnings": [],
    }
    payload.update(overrides)
    return FileInspection.from_dict(payload)


@pytest.mark.parametrize("locale", _LOCALES)
def test_every_locale_has_complete_file_admission_tables(locale: str) -> None:
    table = _load(locale)
    main_window = table.get("main_window")
    admission = table.get("file_admission")

    assert isinstance(main_window, dict)
    assert isinstance(admission, dict)
    assert set(main_window) >= _MAIN_WINDOW_KEYS
    assert set(admission) >= set(_CODE_TO_KEY.values()) | {"actual_format"}
    for key in (*_MAIN_WINDOW_KEYS, *_CODE_TO_KEY.values()):
        source = main_window if key in _MAIN_WINDOW_KEYS else admission
        assert str(source[key]).strip(), f"{locale}: empty translation for {key}"


@pytest.mark.parametrize("locale", _LOCALES)
def test_admission_templates_have_stable_format_parameters(locale: str) -> None:
    admission = _load(locale)["file_admission"]
    assert isinstance(admission, dict)
    formatter = string.Formatter()

    for key in _CODE_TO_KEY.values():
        fields = {
            field_name
            for _literal, field_name, _format_spec, _conversion in formatter.parse(str(admission[key]))
            if field_name
        }
        expected = set() if key == "empty" else {"declared_format", "detected_format"}
        assert fields == expected, f"{locale}.{key}: {fields} != {expected}"


@pytest.mark.parametrize("locale", _LOCALES)
def test_all_stable_codes_render_without_leaking_placeholders(locale: str) -> None:
    original_locale = get_locale()
    try:
        set_locale(locale)
        for code, short_key in _CODE_TO_KEY.items():
            assert file_admission_message_key(code) == f"file_admission.{short_key}"
            rendered = render_file_admission_code(
                code,
                declared_format="docx",
                detected_format="pdf",
                fallback="UNTRANSLATED",
            )
            assert rendered
            assert rendered != "UNTRANSLATED"
            assert "{" not in rendered and "}" not in rendered
            if code != "FILE_EMPTY":
                assert "DOCX" in rendered
                assert "PDF" in rendered
    finally:
        set_locale(original_locale)


def test_message_renderer_prefers_specific_warning_and_preserves_unknown_fallback() -> None:
    rendered = render_file_admission_message(
        warning_code="FILE_FORMAT_CROSS_FAMILY_MISMATCH",
        reason_code="FILE_FORMAT_CONFIRMATION_REQUIRED",
        declared_format="docx",
        detected_format="pdf",
        fallback="fallback",
    )
    assert "DOCX" in rendered
    assert "PDF" in rendered

    assert (
        render_file_admission_message(
            warning_code="FUTURE_WARNING",
            reason_code="FUTURE_REASON",
            declared_format="docx",
            detected_format="pdf",
            fallback="future fallback",
        )
        == "future fallback"
    )


def test_inspection_renderer_localizes_warning_and_can_prefer_blocking_reason() -> None:
    original_locale = get_locale()
    try:
        set_locale("zh_CN")
        warning = _inspection(
            declared_format="docx",
            detected_format="pdf",
            warning_code="FILE_FORMAT_CROSS_FAMILY_MISMATCH",
            warning_message="CORE ENGLISH WARNING",
            reason_code="FILE_FORMAT_CONFIRMATION_REQUIRED",
            reason_message="CORE ENGLISH REASON",
            warnings=(),
        )
        rendered = render_file_inspection_message(warning)
        assert "DOCX" in rendered and "PDF" in rendered
        assert "CORE ENGLISH" not in rendered

        blocked = _inspection(
            declared_format="docx",
            detected_format="zip",
            warning_code="",
            warning_message="",
            reason_code="FILE_CONTAINER_INVALID",
            reason_message="CORE ENGLISH BLOCK",
            warnings=(),
        )
        rendered_block = render_file_inspection_message(blocked, prefer_reason=True)
        assert "DOCX" in rendered_block and "ZIP" in rendered_block
        assert "CORE ENGLISH" not in rendered_block
    finally:
        set_locale(original_locale)


def test_inspection_renderer_preserves_additional_unmapped_diagnostics() -> None:
    inspection = _inspection(
        declared_format="docx",
        detected_format="docx",
        warning_code="FILE_FORMAT_SAME_FAMILY_MISMATCH",
        warning_message="combined core fallback",
        reason_code="",
        reason_message="",
        warnings=(
            {
                "code": "FILE_FORMAT_SAME_FAMILY_MISMATCH",
                "message": "core mismatch fallback",
            },
            {
                "code": "OOXML_SIGNATURE_VALIDATION_UNAVAILABLE",
                "message": "Signature verification is unavailable.",
            },
        ),
    )

    rendered = render_file_inspection_message(inspection)

    assert "DOCX" in rendered
    assert "[OOXML_SIGNATURE_VALIDATION_UNAVAILABLE] Signature verification is unavailable." in rendered


@pytest.mark.parametrize("locale", _LOCALES)
def test_compact_format_notice_uses_detected_format_only_for_mismatch_codes(locale: str) -> None:
    original_locale = get_locale()
    try:
        set_locale(locale)
        inspection = _inspection(
            declared_format="docx",
            detected_format="doc",
            warning_code="FILE_FORMAT_SAME_FAMILY_MISMATCH",
            warning_message="long diagnostic",
        )
        ref = FileRef(
            path="sample.docx",
            format="doc",
            category="document",
            warning_message="long diagnostic",
            metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
        )
        notice = render_file_format_notice(ref)
        assert "DOC" in notice
        assert "DOCX" not in notice
        assert "{" not in notice and "}" not in notice

        exact = _inspection(
            declared_format="docx",
            detected_format="docx",
            warning_code="",
            warning_message="",
        )
        ref.metadata[FILE_INSPECTION_METADATA_KEY] = exact.to_dict()
        assert render_file_format_notice(ref) == ""
    finally:
        set_locale(original_locale)
