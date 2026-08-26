"""Tests for shared OFD patch and OCR helper modules (F-I2a-003, F-I2b-019).

Covers:
- ``docwen_core.text`` — language constants, locale mapping, model mapping,
  and ``resolve_ocr_language``.
- ``docwen_core.ofd`` — ``apply_easyofd_patches`` is importable and
  idempotent.
"""

from __future__ import annotations

import base64
import io
import tempfile
import threading
import time
import warnings
import zipfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


# ═══════════════════════════════════════════════════════════════════════════
# docwen_core.text — constants
# ═══════════════════════════════════════════════════════════════════════════


class TestOcrLanguageConstants:
    """OCR language constant values."""

    def test_constants_are_strings(self) -> None:
        from docwen_core.text import (
            OCR_LANGUAGE_AUTO,
            OCR_LANGUAGE_CHINESE,
            OCR_LANGUAGE_CYRILLIC,
            OCR_LANGUAGE_ENGLISH,
            OCR_LANGUAGE_JAPANESE,
            OCR_LANGUAGE_KOREAN,
            OCR_LANGUAGE_LATIN,
        )

        assert OCR_LANGUAGE_AUTO == "auto"
        assert OCR_LANGUAGE_CHINESE == "chinese"
        assert OCR_LANGUAGE_ENGLISH == "english"
        assert OCR_LANGUAGE_JAPANESE == "japanese"
        assert OCR_LANGUAGE_KOREAN == "korean"
        assert OCR_LANGUAGE_LATIN == "latin"
        assert OCR_LANGUAGE_CYRILLIC == "cyrillic"

    def test_all_languages_have_model_entries(self) -> None:
        from docwen_core.text import (
            OCR_LANGUAGE_CHINESE,
            OCR_LANGUAGE_CHINESE_CHT,
            OCR_LANGUAGE_CYRILLIC,
            OCR_LANGUAGE_ENGLISH,
            OCR_LANGUAGE_JAPANESE,
            OCR_LANGUAGE_KOREAN,
            OCR_LANGUAGE_LATIN,
            OCR_LANGUAGE_MODELS,
        )

        for lang in (
            OCR_LANGUAGE_CHINESE,
            OCR_LANGUAGE_CHINESE_CHT,
            OCR_LANGUAGE_ENGLISH,
            OCR_LANGUAGE_JAPANESE,
            OCR_LANGUAGE_KOREAN,
            OCR_LANGUAGE_LATIN,
            OCR_LANGUAGE_CYRILLIC,
        ):
            assert lang in OCR_LANGUAGE_MODELS, f"{lang} missing from OCR_LANGUAGE_MODELS"
            entry = OCR_LANGUAGE_MODELS[lang]
            assert "det" in entry
            assert "rec" in entry
            assert "cls" in entry
            assert entry["det"].endswith(".onnx")
            assert entry["rec"].endswith(".onnx")
            assert entry["cls"].endswith(".onnx")


class TestLocaleToOcrLanguage:
    """Locale → OCR language mapping coverage."""

    def test_all_locales_map_to_valid_languages(self) -> None:
        from docwen_core.text import LOCALE_TO_OCR_LANGUAGE, OCR_LANGUAGE_MODELS

        for locale, lang in LOCALE_TO_OCR_LANGUAGE.items():
            assert lang in OCR_LANGUAGE_MODELS, f"locale {locale} → {lang} not in OCR_LANGUAGE_MODELS"

    def test_known_locales(self) -> None:
        from docwen_core.text import (
            LOCALE_TO_OCR_LANGUAGE,
            OCR_LANGUAGE_CHINESE,
            OCR_LANGUAGE_CHINESE_CHT,
            OCR_LANGUAGE_CYRILLIC,
            OCR_LANGUAGE_ENGLISH,
            OCR_LANGUAGE_JAPANESE,
            OCR_LANGUAGE_KOREAN,
            OCR_LANGUAGE_LATIN,
        )

        assert LOCALE_TO_OCR_LANGUAGE["zh_CN"] == OCR_LANGUAGE_CHINESE
        assert LOCALE_TO_OCR_LANGUAGE["zh_TW"] == OCR_LANGUAGE_CHINESE_CHT
        assert LOCALE_TO_OCR_LANGUAGE["en_US"] == OCR_LANGUAGE_ENGLISH
        assert LOCALE_TO_OCR_LANGUAGE["ja_JP"] == OCR_LANGUAGE_JAPANESE
        assert LOCALE_TO_OCR_LANGUAGE["ko_KR"] == OCR_LANGUAGE_KOREAN
        assert LOCALE_TO_OCR_LANGUAGE["ru_RU"] == OCR_LANGUAGE_CYRILLIC
        # Latin group
        for loc in ("de_DE", "fr_FR", "pt_BR", "es_ES", "vi_VN"):
            assert LOCALE_TO_OCR_LANGUAGE[loc] == OCR_LANGUAGE_LATIN


# ═══════════════════════════════════════════════════════════════════════════
# resolve_ocr_language
# ═══════════════════════════════════════════════════════════════════════════


class TestResolveOcrLanguage:
    """Tests for ``resolve_ocr_language``."""

    def test_explicit_language_passthrough(self) -> None:
        from docwen_core.text import resolve_ocr_language

        assert resolve_ocr_language("chinese") == "chinese"
        assert resolve_ocr_language("japanese") == "japanese"
        assert resolve_ocr_language("latin") == "latin"
        assert resolve_ocr_language("cyrillic") == "cyrillic"

    def test_auto_resolves_from_locale(self) -> None:
        from docwen_core.text import resolve_ocr_language

        assert resolve_ocr_language("auto", "ja_JP") == "japanese"
        assert resolve_ocr_language("auto", "ko_KR") == "korean"
        assert resolve_ocr_language("auto", "de_DE") == "latin"
        assert resolve_ocr_language("auto", "ru_RU") == "cyrillic"

    def test_auto_defaults_to_chinese_for_unknown_locale(self) -> None:
        from docwen_core.text import resolve_ocr_language

        assert resolve_ocr_language("auto", "xx_XX") == "chinese"
        assert resolve_ocr_language("auto") == "chinese"

    def test_auto_with_chinese_locale(self) -> None:
        from docwen_core.text import resolve_ocr_language

        assert resolve_ocr_language("auto", "zh_CN") == "chinese"
        assert resolve_ocr_language("auto", "zh_TW") == "chinese_cht"


# ═══════════════════════════════════════════════════════════════════════════
# docwen_core.ofd — apply_easyofd_patches
# ═══════════════════════════════════════════════════════════════════════════


class TestOfdPatches:
    """Tests for the OFD monkey-patch module (F-I2a-003)."""

    def test_apply_easyofd_patches_is_importable(self) -> None:
        """The function must be importable and callable."""
        from docwen_core.ofd import apply_easyofd_patches

        assert callable(apply_easyofd_patches)

    def test_patch_guard_prevents_double_application(self) -> None:
        """Calling apply_easyofd_patches multiple times is idempotent."""
        from docwen_core.ofd import apply_easyofd_patches

        apply_easyofd_patches()
        apply_easyofd_patches()

    def test_patch_coordinator_silences_easyofd_loguru_namespace(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_core.ofd as ofd

        disabled: list[str] = []
        monkeypatch.setattr(ofd, "silence_easyofd_import_logging", lambda: disabled.append("easyofd"))

        ofd.apply_easyofd_patches()

        assert disabled == ["easyofd"]

    def test_import_logging_boundary_does_not_import_easyofd(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import builtins

        import docwen_core.ofd as ofd

        real_import = builtins.__import__
        imports: list[str] = []

        def recording_import(name, *args, **kwargs):
            imports.append(name)
            return real_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", recording_import)
        ofd.silence_easyofd_import_logging()

        assert not any(name == "easyofd" or name.startswith("easyofd.") for name in imports)

    def test_import_boundary_suppresses_dependency_syntax_warnings(self) -> None:
        from docwen_core.ofd import easyofd_import_boundary

        with warnings.catch_warnings(record=True) as recorded:
            warnings.simplefilter("always")
            with easyofd_import_boundary():
                warnings.warn("upstream compile warning", SyntaxWarning, stacklevel=1)

        assert recorded == []

    def test_stdout_patch_shadows_only_easyofd_modules(self) -> None:
        import builtins

        import docwen_core.ofd as ofd

        assert ofd._patch_easyofd_stdout()
        for module_name in ofd._EASYOFD_STDOUT_MODULES:
            module = __import__(module_name, fromlist=["*"])
            assert module.print is ofd._discard_easyofd_print
        assert builtins.print is not ofd._discard_easyofd_print

    def test_patch_coordinator_fails_closed_and_remains_retryable(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_core.ofd as ofd

        monkeypatch.setattr(ofd, "_patches_applied", False)
        results = {
            "_patch_easyofd_stdout": True,
            "_patch_draw_annotation": True,
            "_patch_draw_pdf_page_scale": False,
            "_patch_content_clip_boundary": True,
            "_patch_fileread": True,
        }
        for name in tuple(results):
            monkeypatch.setattr(ofd, name, lambda name=name: results[name])

        with pytest.raises(RuntimeError, match="page_scale"):
            ofd.apply_easyofd_patches()
        assert ofd._patches_applied is False

        results["_patch_draw_pdf_page_scale"] = True
        ofd.apply_easyofd_patches()
        assert ofd._patches_applied is True

    def test_patch_coordinator_serializes_concurrent_first_use(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_core.ofd as ofd

        monkeypatch.setattr(ofd, "_patches_applied", False)
        calls: list[str] = []
        call_lock = threading.Lock()

        def record(name: str) -> bool:
            with call_lock:
                calls.append(name)
            time.sleep(0.01)
            return True

        for name in (
            "_patch_draw_annotation",
            "_patch_draw_pdf_page_scale",
            "_patch_content_clip_boundary",
            "_patch_fileread",
        ):
            monkeypatch.setattr(ofd, name, lambda name=name: record(name))

        errors: list[BaseException] = []

        def apply() -> None:
            try:
                ofd.apply_easyofd_patches()
            except BaseException as exc:  # pragma: no cover - assertion carrier
                errors.append(exc)

        threads = [threading.Thread(target=apply) for _ in range(3)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()

        assert errors == []
        assert sorted(calls) == sorted(
            [
                "_patch_draw_annotation",
                "_patch_draw_pdf_page_scale",
                "_patch_content_clip_boundary",
                "_patch_fileread",
            ]
        )

    def test_submodules_are_importable(self) -> None:
        """The individual easyofd patch helpers are importable."""
        from docwen_core.ofd import (
            _patch_content_clip_boundary,
            _patch_draw_annotation,
            _patch_draw_pdf_page_scale,
            _patch_fileread,
        )

        assert callable(_patch_draw_annotation)
        assert callable(_patch_draw_pdf_page_scale)
        assert callable(_patch_content_clip_boundary)
        assert callable(_patch_fileread)

    def test_fileread_owns_cleanup_under_dotted_temp_parent(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """A dotted TEMP path must never become an ancestor cleanup target."""
        from easyofd.parser_ofd.file_deal import FileRead

        from docwen_core.ofd import _patch_fileread

        candidate_root = tmp_path / "candidate-root"
        dotted_parent = candidate_root / ".pytest-runtime" / "system-temp"
        dotted_parent.mkdir(parents=True)
        sentinel = candidate_root / "sentinel.txt"
        sentinel.write_text("retain\n", encoding="utf-8")
        monkeypatch.setattr(tempfile, "tempdir", str(dotted_parent))
        _patch_fileread()

        payload = io.BytesIO()
        with zipfile.ZipFile(payload, "w") as archive:
            archive.writestr("OFD.xml", '<ofd:OFD xmlns:ofd="urn:ofd"/>')
        reader = FileRead(base64.b64encode(payload.getvalue()).decode("ascii"))
        owner = Path(reader.__dict__["_docwen_scratch_root"])

        file_tree = reader()

        assert Path(file_tree["root_doc"]).parts[-2:] == ("unpacked", "OFD.xml")
        assert sentinel.read_text(encoding="utf-8") == "retain\n"
        assert not owner.exists()
        assert list(dotted_parent.iterdir()) == []

    def test_fileread_cleans_owned_scratch_after_malformed_archive(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from easyofd.parser_ofd.file_deal import FileRead

        from docwen_core.ofd import _patch_fileread

        candidate_root = tmp_path / "candidate-root"
        dotted_parent = candidate_root / ".pytest-runtime"
        dotted_parent.mkdir(parents=True)
        sentinel = candidate_root / "sentinel.txt"
        sentinel.write_text("retain\n", encoding="utf-8")
        monkeypatch.setattr(tempfile, "tempdir", str(dotted_parent))
        _patch_fileread()
        reader = FileRead(base64.b64encode(b"not-an-ofd-archive").decode("ascii"))
        owner = Path(reader.__dict__["_docwen_scratch_root"])

        with pytest.raises(zipfile.BadZipFile):
            reader()

        assert sentinel.read_text(encoding="utf-8") == "retain\n"
        assert not owner.exists()
        assert list(dotted_parent.iterdir()) == []

    def test_fileread_rejects_archive_member_path_escape(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from easyofd.parser_ofd.file_deal import FileRead

        from docwen_core.ofd import _patch_fileread

        candidate_root = tmp_path / "candidate-root"
        dotted_parent = candidate_root / ".pytest-runtime"
        dotted_parent.mkdir(parents=True)
        sentinel = candidate_root / "sentinel.txt"
        sentinel.write_text("retain\n", encoding="utf-8")
        monkeypatch.setattr(tempfile, "tempdir", str(dotted_parent))
        _patch_fileread()
        payload = io.BytesIO()
        with zipfile.ZipFile(payload, "w") as archive:
            archive.writestr("../escaped.txt", "escape")
        reader = FileRead(base64.b64encode(payload.getvalue()).decode("ascii"))
        owner = Path(reader.__dict__["_docwen_scratch_root"])

        with pytest.raises(ValueError, match="unsafe OFD archive member path"):
            reader()

        assert sentinel.read_text(encoding="utf-8") == "retain\n"
        assert not (candidate_root / "escaped.txt").exists()
        assert not owner.exists()
        assert list(dotted_parent.iterdir()) == []

    def test_draw_pdf_uses_pdf_points_per_millimetre(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """A 210 x 297 mm OFD page must map to standard A4 PDF points."""
        import easyofd.draw.draw_pdf as draw_pdf_module

        from docwen_core.ofd import _patch_draw_pdf_page_scale

        _patch_draw_pdf_page_scale()
        monkeypatch.setattr(draw_pdf_module, "FontTool", lambda: object())

        draw_pdf = draw_pdf_module.DrawPDF([{"pdf_name": "probe"}])

        assert pytest.approx(595.2756 / 210, abs=0.00001) == draw_pdf.OP
        assert pytest.approx(841.8898 / 297, abs=0.00001) == draw_pdf.OP

    def test_unbounded_abbreviated_clip_does_not_crash_text_parser(self) -> None:
        """An abbreviated clip path without @Boundary must remain parseable."""
        from easyofd.parser_ofd.file_content_parser import ContentFileParser

        from docwen_core.ofd import _patch_content_clip_boundary

        _patch_content_clip_boundary()
        parser = ContentFileParser({"ofd:Page": {}})
        row = {
            "@ID": "14",
            "@Boundary": "169.902 245.177 3.7977 4.9565",
            "@Font": "6",
            "@Size": "4.9565",
            "ofd:Clips": {
                "ofd:Clip": {
                    "ofd:Area": {
                        "ofd:Path": {
                            "@Stroke": "false",
                            "@Fill": "true",
                            "ofd:AbbreviatedData": "M 0 0 L 1 1 C",
                        }
                    }
                }
            },
        }

        result = parser.fetch_cell_info(row, {"#text": "-"})

        assert result["text"] == "-"
        assert "clips_pos" not in result
        assert "ofd:Clips" in row

    def test_bounded_clip_retains_upstream_position_parsing(self) -> None:
        """The compatibility patch must not alter bounded clip positions."""
        from easyofd.parser_ofd.file_content_parser import ContentFileParser

        from docwen_core.ofd import _patch_content_clip_boundary

        _patch_content_clip_boundary()
        parser = ContentFileParser({"ofd:Page": {}})
        row = {
            "@ID": "15",
            "@Boundary": "1 2 3 4",
            "@Font": "6",
            "@Size": "5",
            "ofd:Clips": {"ofd:Clip": {"ofd:Area": {"ofd:Path": {"@Boundary": "5 6 7 8"}}}},
        }

        result = parser.fetch_cell_info(row, {"#text": "bounded"})

        assert result["clips_pos"] == [5.0, 6.0, 7.0, 8.0]
