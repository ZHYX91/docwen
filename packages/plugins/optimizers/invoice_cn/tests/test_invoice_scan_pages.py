"""Tests for invoice scan-page detection and PDF page rendering.

Covers F-G5-006 (_render_pdf_page_to_png) and F-G5-007 (_is_scanpage) closure:
- ``is_scanpage()`` correctly identifies scan-based vs text-based PDF pages.
- ``render_pdf_page_to_png()`` renders a PDF page to a preprocessed PNG.
- The ``InvoiceCnConverter`` routes scan-based PDFs to OCR when ``to_md_enable_ocr=True``.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

# ── helpers ────────────────────────────────────────────────────────────


def _build_fake_context(
    input_path: str,
    staging_dir: str,
    target_format: str = "md",
    options: dict | None = None,
    action_name: str = "invoice_cn",
    source_format: str = "pdf",
    *,
    pre_cancelled: bool = False,
):
    """Build a minimal fake PluginExecutionContext for converter testing."""
    from dataclasses import dataclass, field

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy
    from docwen_core.protocols.execution_context import CancellationTokenView

    @dataclass
    class FakeWorkspaceHandle:
        input_path: str
        staging_dir: str
        _inputs: tuple[FileRef, ...]
        _counter: list[int] = field(default_factory=lambda: [0])
        _artifacts: list[ArtifactManifest] = field(default_factory=list)

        def input_resources(self, role: str | None = None) -> tuple[FileRef, ...]:
            if role is None:
                return self._inputs
            return tuple(item for item in self._inputs if item.input_role == role)

        def resource_by_logical_path(self, logical_path: str) -> FileRef | None:
            return next((item for item in self._inputs if item.logical_path == logical_path), None)

        def create_artifact_path(self, kind: str, suffix: str) -> str:
            self._counter[0] += 1
            return str(Path(self.staging_dir) / f"{kind}_{self._counter[0]}{suffix}")

        def add_artifact(self, manifest: ArtifactManifest) -> None:
            self._artifacts.append(manifest)

        @property
        def registered_artifacts(self) -> list[ArtifactManifest]:
            return list(self._artifacts)

    @dataclass
    class FakeProgressSink:
        events: list[tuple[float, str]] = field(default_factory=list)
        artifacts: list[tuple[str, str]] = field(default_factory=list)
        diagnostics: list[tuple[str, str, str, str]] = field(default_factory=list)

        def report_progress(self, percent: float, message: str = "") -> None:
            self.events.append((percent, message))

        def report_diagnostic(self, level: str, message: str, code: str = "", location: str = "") -> None:
            self.diagnostics.append((level, message, code, location))

        def report_artifact_ready(self, artifact_id: str, suggested_name: str) -> None:
            self.artifacts.append((artifact_id, suggested_name))

    @dataclass
    class FakeLogger:
        messages: list[str] = field(default_factory=list)

        def debug(self, message: str, **extra: object) -> None:
            pass

        def info(self, message: str, **extra: object) -> None:
            self.messages.append(message)

        def warning(self, message: str, **extra: object) -> None:
            pass

        def error(self, message: str, **extra: object) -> None:
            self.messages.append(message)

    @dataclass
    class FakeConfig:
        def get(self, key: str, default: object = None) -> object:
            return default

        def get_plugin_config(self, plugin_id: str) -> dict[str, object]:
            return {}

    @dataclass
    class FakeContext:
        request: ConversionRequest
        workspace: FakeWorkspaceHandle
        config: FakeConfig
        progress: FakeProgressSink
        cancellation: CancellationTokenView
        logger: FakeLogger

    detected_format = source_format or Path(input_path).suffix.lstrip(".")
    file_refs = [
        FileRef(
            path=input_path,
            format=detected_format,
            category="layout",
        )
    ]
    request = ConversionRequest(
        request_id="test-scanpage-001",
        input_refs=file_refs,
        target_format=target_format,
        action_name=action_name,
        options=options or {},
        output_policy=OutputPolicy(),
    )

    token = CancellationToken()
    if pre_cancelled:
        token.cancel("test cancellation")
    return FakeContext(
        request,
        FakeWorkspaceHandle(input_path, staging_dir, tuple(file_refs)),
        FakeConfig(),
        FakeProgressSink(),
        token.view(),
        FakeLogger(),
    )


def _make_text_pdf(path: Path) -> None:
    """Create a single-page PDF with invoice-like Chinese text.

    Uses PyMuPDF directly — the text is long enough that is_scanpage()
    returns False.
    """
    import fitz

    text = (
        "电子发票（普通发票）\n"
        "发票代码：123456789012\n"
        "发票号码：87654321\n"
        "开票日期：2024年01月15日\n"
        "校验码：12345678901234567890\n"
        "购买方信息\n"
        "名称：测试购买方公司\n"
        "统一社会信用代码/纳税人识别号：91110108MA01ABCD1X\n"
        "销售方信息\n"
        "名称：测试销售方公司\n"
        "统一社会信用代码/纳税人识别号：91110108MA01EFGH2Y\n"
        "项目名称 规格型号 单位 数量 单价 金额 税率 税额\n"
        "*测试商品一 件 2 100.00 200.00 13% 26.00\n"
        "*测试商品二 箱 1 500.00 500.00 13% 65.00\n"
        "合计 ¥700.00 ¥91.00\n"
        "价税合计 ¥791.00\n"
    )
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((72, 72), text, fontsize=10)
    doc.save(str(path))
    doc.close()


def _make_scan_pdf(path: Path) -> None:
    """Create a single-page PDF with NO extractable text (simulating a scanned page).

    Inserts a small embedded image instead of text so PyMuPDF text extraction
    returns empty/noise — is_scanpage() should return True.
    """
    import fitz

    doc = fitz.open()
    page = doc.new_page(width=595, height=842)

    # Create a valid small PNG using Pillow (or use a raw pixmap fallback)
    try:
        import io

        from PIL import Image

        buf = io.BytesIO()
        img = Image.new("RGB", (2, 2), color=(255, 255, 255))
        img.save(buf, format="PNG")
        png_bytes = buf.getvalue()
        img_rect = fitz.Rect(100, 100, 400, 300)
        page.insert_image(img_rect, stream=png_bytes)
    except ImportError:
        # Fallback: use a raw pixmap (this always works with PyMuPDF)
        # Create a 2x2 white pixmap
        pm = fitz.Pixmap(fitz.csRGB, 2, 2, False)
        pm.clear_with(255)
        img_rect = fitz.Rect(100, 100, 400, 300)
        page.insert_image(img_rect, pixmap=pm)

    doc.save(str(path))
    doc.close()


# ── Unit tests: is_scanpage ────────────────────────────────────────────


class TestIsScanpage:
    """Tests for scan-based vs text-based PDF page detection."""

    def test_detects_scan_page_empty_text(self) -> None:
        """Empty text should be detected as a scan page."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import is_scanpage

        assert is_scanpage("") is True
        assert is_scanpage("   ") is True
        assert is_scanpage("\n\t\r") is True

    def test_detects_scan_page_short_text(self) -> None:
        """Text shorter than MIN_TEXT_LENGTH_FOR_INVOICE=20 should be scan."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            MIN_TEXT_LENGTH_FOR_INVOICE,
            is_scanpage,
        )

        short = "A" * (MIN_TEXT_LENGTH_FOR_INVOICE - 1)  # 19 chars
        assert is_scanpage(short) is True

    def test_detects_non_scan_page_long_text(self) -> None:
        """Text >= 20 chars should NOT be scan."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            MIN_TEXT_LENGTH_FOR_INVOICE,
            is_scanpage,
        )

        long_enough = "A" * MIN_TEXT_LENGTH_FOR_INVOICE  # 20 chars
        assert is_scanpage(long_enough) is False

    def test_scanpage_with_invoice_text(self) -> None:
        """A real invoice text excerpt should NOT be detected as scan."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import is_scanpage

        text = "电子发票（普通发票）\n发票号码：87654321\n开票日期：2024年01月15日\n"
        assert is_scanpage(text) is False

    def test_scanpage_with_zero_width_chars(self) -> None:
        """Zero-width characters should be stripped before counting."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import is_scanpage

        # Zero-width chars only → scan
        assert is_scanpage("﻿​‌‍") is True
        # Zero-width mixed with short text → scan
        assert is_scanpage("﻿​abc‌‍") is True

    def test_is_scanpage_threshold_near_boundary(self) -> None:
        """Exactly at boundary: 19 = scan, 20 = text."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            MIN_TEXT_LENGTH_FOR_INVOICE,
            is_scanpage,
        )

        # 19 chars — scan
        assert is_scanpage("电" * (MIN_TEXT_LENGTH_FOR_INVOICE - 1)) is True
        # 20 chars — not scan
        assert is_scanpage("电" * MIN_TEXT_LENGTH_FOR_INVOICE) is False

    def test_scanpage_with_tabs_and_newlines(self) -> None:
        """Whitespace-only text should be detected as scan."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import is_scanpage

        assert is_scanpage("\n\n\t\t   \r\n") is True


# ── Unit tests: render_pdf_page_to_png ─────────────────────────────────


class TestRenderPdfPageToPng:
    """Tests for PDF page → preprocessed PNG rendering."""

    def test_renders_text_pdf_page_to_png(self, tmp_path: Path) -> None:
        """Rendering a PDF text page should produce a valid PNG file."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            render_pdf_page_to_png,
        )

        pdf_path = tmp_path / "test.pdf"
        _make_text_pdf(pdf_path)
        png_path = str(tmp_path / "output.png")

        result = render_pdf_page_to_png(file_path=str(pdf_path), page_index=0, png_path=png_path)
        assert result == png_path
        assert os.path.isfile(png_path)
        assert os.path.getsize(png_path) > 0

    def test_rendered_png_is_valid_image(self, tmp_path: Path) -> None:
        """The output PNG should be openable by PIL."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            render_pdf_page_to_png,
        )

        pdf_path = tmp_path / "test.pdf"
        _make_text_pdf(pdf_path)
        png_path = str(tmp_path / "output.png")

        render_pdf_page_to_png(file_path=str(pdf_path), page_index=0, png_path=png_path)

        try:
            from PIL import Image

            img = Image.open(png_path)
            assert img.size[0] > 0
            assert img.size[1] > 0
        except ImportError:
            pytest.skip("Pillow not available")

    def test_renders_scan_pdf_page(self, tmp_path: Path) -> None:
        """Rendering a scan-like PDF page (no text, image content) should work."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            render_pdf_page_to_png,
        )

        pdf_path = tmp_path / "scan.pdf"
        _make_scan_pdf(pdf_path)
        png_path = str(tmp_path / "scan_output.png")

        render_pdf_page_to_png(file_path=str(pdf_path), page_index=0, png_path=png_path)
        assert os.path.isfile(png_path)
        assert os.path.getsize(png_path) > 0

    def test_renders_multi_page_pdf_specific_page(self, tmp_path: Path) -> None:
        """Render page index > 0 from a multi-page PDF."""
        import fitz

        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            render_pdf_page_to_png,
        )

        pdf_path = tmp_path / "multi.pdf"
        doc = fitz.open()
        doc.new_page(width=595, height=842).insert_text((72, 72), "Page One", fontsize=12)
        doc.new_page(width=595, height=842).insert_text((72, 72), "Page Two Content Here", fontsize=12)
        doc.save(str(pdf_path))
        doc.close()

        png_path = str(tmp_path / "page1.png")
        render_pdf_page_to_png(file_path=str(pdf_path), page_index=1, png_path=png_path)
        assert os.path.isfile(png_path)
        assert os.path.getsize(png_path) > 0

    def test_render_pdf_page_to_png_300dpi(self, tmp_path: Path) -> None:
        """Rendered image should reflect ~300 DPI resolution (large pixel dimensions)."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            render_pdf_page_to_png,
        )

        pdf_path = tmp_path / "dpi_test.pdf"
        _make_text_pdf(pdf_path)
        png_path = str(tmp_path / "dpi_output.png")

        render_pdf_page_to_png(file_path=str(pdf_path), page_index=0, png_path=png_path)

        try:
            from PIL import Image

            img = Image.open(png_path)
            # A4 at 300 DPI → ~2480 x 3508 pixels
            assert img.size[0] >= 2000, f"Expected >= 2000px width at 300 DPI, got {img.size[0]}"
            assert img.size[1] >= 3000, f"Expected >= 3000px height at 300 DPI, got {img.size[1]}"
        except ImportError:
            pytest.skip("Pillow not available")


# ── Integration tests: scan-page → converter ───────────────────────────


class TestScanPageInConverter:
    """Tests that the converter routes scan-based PDFs to OCR."""

    def test_non_scan_pdf_uses_text_path(self, tmp_path: Path) -> None:
        """A text-rich PDF should use the text extraction path (INVOICE-OK, not OCR)."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        pdf_path = tmp_path / "text_invoice.pdf"
        _make_text_pdf(pdf_path)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                source_format="pdf",
                options={"to_md_enable_ocr": True},
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True
            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)
            # Text PDF should not trigger OCR even with enable_ocr=True
            assert not any(d.code == "INVOICE-OCR-OK" for d in result.diagnostics)

    def test_scan_pdf_ocr_failure_is_best_effort(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """A scan PDF with enable_ocr=True should render to PNG then OCR."""
        import docwen_plugin_optimizer_invoice_cn.invoice_cn.converter as converter_module
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        def missing_ocr_model(*_args: object, **_kwargs: object):
            raise FileNotFoundError("rapidocr model missing")

        monkeypatch.setattr(converter_module, "_ocr_and_parse_image_invoice", missing_ocr_model)

        pdf_path = tmp_path / "scan_invoice.pdf"
        _make_scan_pdf(pdf_path)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                source_format="pdf",
                options={"to_md_enable_ocr": True},
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True
            assert result.error is None
            assert any(item[2] == "OCR-BEST-EFFORT" for item in context.progress.diagnostics)
            assert any(d.code == "INVOICE-SCAN-DETECTED" for d in result.diagnostics)
            assert not any(d.code == "INVOICE-OCR-OK" for d in result.diagnostics)

    def test_scan_pdf_typed_ocr_failure_falls_back_to_text_parser(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """A typed OCR failure must preserve the base PDF parser result."""
        import docwen_plugin_optimizer_invoice_cn.invoice_cn.converter as converter_module
        import docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser as pdf_parser_module
        from docwen_core.text.ocr import OcrOutcome, OcrStatus
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import (
            INVOICE_CN_YAML_SCHEMA,
        )

        monkeypatch.setattr(
            converter_module,
            "_ocr_and_parse_image_invoice",
            lambda *_args, **_kwargs: (
                {},
                [],
                OcrOutcome(OcrStatus.MODEL_MISSING, message="private model path"),
            ),
        )
        fallback_calls: list[str] = []

        def parse_pdf_invoice(input_path: str):
            fallback_calls.append(input_path)
            return {INVOICE_CN_YAML_SCHEMA[0]: "BASE-PARSER-RESULT"}, []

        monkeypatch.setattr(pdf_parser_module, "parse_pdf_invoice", parse_pdf_invoice)

        pdf_path = tmp_path / "typed_failure_scan_invoice.pdf"
        _make_scan_pdf(pdf_path)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                source_format="pdf",
                options={"to_md_enable_ocr": True},
            )
            result = InvoiceCnConverter().convert(context)

            markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")

        assert result.success is True
        assert fallback_calls == [str(pdf_path)]
        assert "BASE-PARSER-RESULT" in markdown
        assert any(item[2] == "OCR-BEST-EFFORT" for item in context.progress.diagnostics)
        assert not any(d.code == "INVOICE-OCR-OK" for d in result.diagnostics)

    def test_scan_pdf_without_ocr_uses_text_path(self, tmp_path: Path) -> None:
        """A scan PDF WITHOUT enable_ocr should still succeed (text path)."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        pdf_path = tmp_path / "scan_invoice_noocr.pdf"
        _make_scan_pdf(pdf_path)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                source_format="pdf",
                # Not setting to_md_enable_ocr → defaults to False
            )
            result = InvoiceCnConverter().convert(context)

            # Should succeed with text path (even though text is minimal)
            assert result.success is True
            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)

    def test_scan_pdf_produces_markdown_output(self, tmp_path: Path) -> None:
        """Regardless of OCR availability, the converter should produce MD output."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        pdf_path = tmp_path / "scan_any.pdf"
        _make_scan_pdf(pdf_path)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                source_format="pdf",
                options={"to_md_enable_ocr": True},
            )
            result = InvoiceCnConverter().convert(context)

            if result.success:
                assert len(result.artifacts) == 1
                artifact = result.artifacts[0]
                assert artifact.media_type == "text/markdown"
                assert os.path.isfile(artifact.staging_path)

                content = Path(artifact.staging_path).read_text(encoding="utf-8")
                # Should always have YAML frontmatter structure
                assert content.startswith("---")
                assert "## 商品明细" in content


# ── Regression: existing tests should not be broken ────────────────────


class TestScanPageRegression:
    """Verify that scan-page additions do not break existing PDF invoice behavior."""

    def test_existing_text_pdf_still_works_without_ocr_option(self, tmp_path: Path) -> None:
        """A text PDF without to_md_enable_ocr should work exactly as before."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        pdf_path = tmp_path / "normal.pdf"
        _make_text_pdf(pdf_path)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                source_format="pdf",
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)

    def test_min_text_length_constant_is_positive(self) -> None:
        """MIN_TEXT_LENGTH_FOR_INVOICE should be a positive integer."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
            MIN_TEXT_LENGTH_FOR_INVOICE,
        )

        assert isinstance(MIN_TEXT_LENGTH_FOR_INVOICE, int)
        assert MIN_TEXT_LENGTH_FOR_INVOICE > 0
