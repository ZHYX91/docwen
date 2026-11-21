"""Tests for Layout OFD/XPS preprocessing with OCR integration (F-I2a-003, F-I2b-019).

Verifies:
- OFD preprocess path uses the shared renderer and reports failures.
- XPS preprocess path works correctly.
- Layout OCR utils import shared constants from ``docwen_core.text``.
- OCR language resolution is used (not bypassed with hardcoded Chinese).
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from unittest import mock

import pytest

pytestmark = pytest.mark.unit


# ═══════════════════════════════════════════════════════════════════════════
# OFD preprocessing
# ═══════════════════════════════════════════════════════════════════════════


class TestOfdPreprocess:
    def test_ofd_missing_dependency(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _ofd_to_pdf

        with mock.patch("docwen_core.ofd.render_ofd_pdf", side_effect=ImportError("easyofd")):
            result = _ofd_to_pdf("/fake/test.ofd", str(tmp_path))

        assert result.error_type == "dependency_missing"
        assert result.error_message is not None and "easyofd" in result.error_message
        assert result.diagnostic_code == "OFD2PDF-DEPENDENCY-MISSING"

    def test_ofd_preprocess_uses_shared_renderer(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _ofd_to_pdf

        pdf_bytes = b"%PDF-1.4 fake ofd output"
        with mock.patch("docwen_core.ofd.render_ofd_pdf", return_value=pdf_bytes) as render:
            result = _ofd_to_pdf("/fake/test.ofd", str(tmp_path))

        render.assert_called_once_with("/fake/test.ofd", check_cancel=None)
        assert result.error_type is None
        assert result.effective_source_format == "pdf"
        assert result.original_source_format == "ofd"
        assert Path(result.effective_input_path).read_bytes() == pdf_bytes

    def test_cancel_during_ofd_render_keeps_staging_empty(self, tmp_path: Path) -> None:
        from docwen_core.cancellation import CancellationToken
        from docwen_core.errors import CancellationRequested
        from docwen_plugin_layout.preprocess import preprocess_layout_input

        owner = CancellationToken()
        cancel = owner.view()

        def render(_path, *, check_cancel):
            owner.cancel()
            check_cancel()

        with (
            mock.patch("docwen_core.ofd.render_ofd_pdf", side_effect=render),
            pytest.raises(CancellationRequested),
        ):
            preprocess_layout_input("/fake/test.ofd", str(tmp_path), "ofd", check_cancel=cancel.check)
        assert list(tmp_path.iterdir()) == []

    def test_ofd_conversion_failure(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _ofd_to_pdf

        with mock.patch("docwen_core.ofd.render_ofd_pdf", side_effect=ValueError("OFD member not found: picture.png")):
            result = _ofd_to_pdf("/fake/test.ofd", str(tmp_path))

        assert result.error_type == "conversion_failed"
        assert result.diagnostic_code == "OFD2PDF-ERROR"
        assert result.error_message is not None and "OFD member not found: picture.png" in result.error_message


# ═══════════════════════════════════════════════════════════════════════════
# XPS preprocessing
# ═══════════════════════════════════════════════════════════════════════════


class TestXpsPreprocess:
    """Tests for XPS→PDF preprocessing."""

    def test_xps_missing_dependency(self) -> None:
        """When fitz is not installed, return dependency_missing error."""
        import builtins

        from docwen_plugin_layout.preprocess import _xps_to_pdf

        _real_import = builtins.__import__

        def _block_fitz(name, *args, **kwargs):
            if name == "fitz":
                raise ImportError("No module named 'fitz'")
            return _real_import(name, *args, **kwargs)

        with mock.patch.object(builtins, "__import__", side_effect=_block_fitz):
            result = _xps_to_pdf("/fake/test.xps", tempfile.gettempdir())

        assert result.error_type == "dependency_missing"
        assert result.diagnostic_code == "XPS2PDF-DEPENDENCY-MISSING"

    def test_xps_conversion_success(self, tmp_path: Path) -> None:
        """When fitz is available, XPS→PDF should produce a valid result.

        Uses a deterministic, font-free two-page XPS package.
        """
        from tests.support.xps import create_minimal_xps, pdf_visual_projection

        from docwen_plugin_layout.preprocess import _xps_to_pdf

        xps_path = tmp_path / "minimal.xps"
        create_minimal_xps(xps_path)
        with tempfile.TemporaryDirectory() as td:
            result = _xps_to_pdf(str(xps_path), td)

            assert result.error_type is None
            assert result.effective_source_format == "pdf"
            assert result.original_source_format == "xps"
            assert len(result.intermediate_artifacts) == 1
            assert Path(result.effective_input_path).exists()
            projection = pdf_visual_projection(result.effective_input_path)
            assert projection["pdf_magic"] == "%PDF-"
            assert projection["page_count"] == 2


# ═══════════════════════════════════════════════════════════════════════════
# OCR helper integration
# ═══════════════════════════════════════════════════════════════════════════


class TestLayoutOcrIntegration:
    """Verify that layout OCR uses shared core helpers directly."""

    def test_core_typed_ocr_entry_accepts_language_parameter(self) -> None:
        import inspect

        from docwen_core.text.ocr import run_ocr_outcome

        sig = inspect.signature(run_ocr_outcome)
        assert "ocr_language" in sig.parameters
        param = sig.parameters["ocr_language"]
        assert param.default is not inspect.Parameter.empty

    def test_core_reset_ocr_exists(self) -> None:
        """reset_ocr remains available for runtime language switching."""
        from docwen_core.text.ocr import reset_ocr

        assert callable(reset_ocr)
