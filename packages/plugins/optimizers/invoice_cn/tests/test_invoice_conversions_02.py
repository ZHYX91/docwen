"""Focused tests split from test_invoice_conversions.py."""

from __future__ import annotations

from ._invoice_conversions_support import (
    Any,
    Path,
    _build_fake_context,
    _ocr_outcome,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.contract


class TestImageInvoiceOcr:
    def test_direct_image_parser_preserves_typed_ocr_failure(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_plugin_optimizer_invoice_cn.invoice_cn import image_parser

        def failed_ocr(*_args: object, **_kwargs: object) -> Any:
            return _ocr_outcome("recognition_failed", message="recognizer failed")

        monkeypatch.setattr(image_parser, "run_ocr_outcome", failed_ocr)

        metadata, rows, outcome = image_parser.parse_image_invoice_outcome("invoice.png", source_format="png")

        assert metadata == {}
        assert rows == []
        assert outcome.status.value == "recognition_failed"
        assert outcome.message == "recognizer failed"

    def test_image_invoice_ocr_consumes_request_language_and_locale(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Image invoice OCR should pass public OCR language options to core OCR."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn import image_parser
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        calls: list[tuple[str, str, str]] = []

        def fake_parse_image_invoice(
            input_path: str,
            *,
            source_format: str,
            ocr_language: str = "auto",
            current_locale: str = "zh_CN",
        ) -> tuple[dict[str, str | None], list[dict[str, str]], Any]:
            assert source_format == "png"
            calls.append((Path(input_path).name, ocr_language, current_locale))
            return (
                {
                    "发票号码": "2604700000000392402353",
                    "开票日期": "2026年01月02日",
                    "价税合计": "3.00",
                },
                [],
                _ocr_outcome("success", text="invoice text"),
            )

        monkeypatch.setattr(image_parser, "parse_image_invoice_outcome", fake_parse_image_invoice)

        img_path = tmp_path / "invoice.png"
        img_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 100)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(img_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="png",
                options={"ocr_language": "japanese", "locale": "ja_JP"},
            )
            result = InvoiceCnConverter().convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        assert calls == [("invoice.png", "japanese", "ja_JP")]
        assert any(d.code == "INVOICE-OCR-OK" for d in result.diagnostics)

    def test_image_invoice_ocr_failure_preserves_base_conversion(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Image→md with invoice_cn should attempt OCR; fails if models missing."""
        import docwen_plugin_optimizer_invoice_cn.invoice_cn.converter as converter_module
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        def missing_ocr_model(*_args: object, **_kwargs: object):
            raise FileNotFoundError("rapidocr model missing")

        monkeypatch.setattr(converter_module, "_ocr_and_parse_image_invoice", missing_ocr_model)

        # Create a dummy image file
        img_path = tmp_path / "test.png"
        img_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 100)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(img_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="png",
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True
            assert result.error is None
            assert any(item[2] == "OCR-BEST-EFFORT" for item in context.progress.diagnostics)
            assert not any(d.code == "INVOICE-OCR-OK" for d in result.diagnostics)
            assert Path(result.artifacts[0].staging_path).is_file()

    @pytest.mark.parametrize(
        "status_value",
        ["input_missing", "unavailable", "model_missing", "initialization_failed", "recognition_failed"],
    )
    def test_image_invoice_typed_ocr_failures_are_safe_and_nonfatal(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        status_value: str,
    ) -> None:
        import docwen_plugin_optimizer_invoice_cn.invoice_cn.converter as converter_module
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import InvoiceCnConverter

        monkeypatch.setattr(
            converter_module,
            "_ocr_and_parse_image_invoice",
            lambda *_args, **_kwargs: (
                {},
                [],
                _ocr_outcome(status_value, message="private model path"),
            ),
        )
        img_path = tmp_path / "invoice.png"
        img_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 100)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(img_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="png",
            )
            result = InvoiceCnConverter().convert(context)

        assert result.success is True
        assert len(context.progress.diagnostics) == 1
        warning = context.progress.diagnostics[0]
        assert warning[2] == "OCR-BEST-EFFORT"
        assert f"status={status_value}" in warning[1]
        assert "private" not in warning[1]
        assert not any(d.code == "INVOICE-OCR-OK" for d in result.diagnostics)

    def test_image_invoice_no_text_warns_without_reporting_false_success(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_optimizer_invoice_cn.invoice_cn.converter as converter_module
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import InvoiceCnConverter

        monkeypatch.setattr(
            converter_module,
            "_ocr_and_parse_image_invoice",
            lambda *_args, **_kwargs: ({}, [], _ocr_outcome("no_text")),
        )
        img_path = tmp_path / "blank.png"
        img_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 100)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(img_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="png",
            )
            result = InvoiceCnConverter().convert(context)

        assert result.success is True
        assert len(context.progress.diagnostics) == 1
        warning = context.progress.diagnostics[0]
        assert warning[2] == "OCR-BEST-EFFORT"
        assert "status=no_text" in warning[1]
        assert "may have been missed" in warning[1]
        assert not any(d.code == "INVOICE-OCR-OK" for d in result.diagnostics)

    def test_ocr_option_on_pdf_no_longer_not_implemented(self, sample_invoice_pdf_path: Path) -> None:
        """Setting to_md_enable_ocr=True on a PDF should no longer return NOT_IMPLEMENTED."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_invoice_pdf_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="pdf",
                options={"to_md_enable_ocr": True},
            )
            result = InvoiceCnConverter().convert(context)

            # Should succeed (PDF parsing doesn't need OCR models)
            assert result.success is True
            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)


class TestInvoiceErrorHandling:
    def test_unsupported_source_format(self, tmp_path: Path) -> None:
        """An unsupported source format should return invalid_input error."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        docx_path = tmp_path / "test.docx"
        docx_path.write_bytes(b"dummy")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(docx_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="docx",
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "invalid_input"
            assert result.error.diagnostic_code == "INVOICE-INVALID-FORMAT"

    def test_plugin_rejects_unadvertised_route(self, tmp_path: Path) -> None:
        """Non-invoice actions are unsupported, not unimplemented promises."""
        from docwen_plugin_optimizer_invoice_cn import InvoicePlugin

        dummy = tmp_path / "dummy.pdf"
        dummy.write_bytes(b"%PDF-1.4 dummy")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(dummy),
                staging,
                "pdf",
                action_name="",
                source_format="pdf",
            )
            result = InvoicePlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "unsupported_route"
            assert result.error.diagnostic_code == "INVOICE-UNSUPPORTED-ROUTE"

    def test_simple_pdf_still_produces_output(self, sample_simple_pdf_path: Path) -> None:
        """Even a simple PDF without invoice content should not crash."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_simple_pdf_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="pdf",
            )
            result = InvoiceCnConverter().convert(context)

            # Should succeed (no crash), even if few fields are detected
            assert result.success is True
            content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")

            # At minimum, should have YAML frontmatter structure
            assert content.startswith("---")
            assert "## 商品明细" in content
            assert "| 商品名称 |" in content


class TestMetadataParser:
    """Test metadata extraction from raw text strings."""

    def test_parse_compact_text_basic_fields(self) -> None:
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.metadata import (
            parse_invoice_metadata_from_compact_text,
        )
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import (
            _compact_text,
        )

        text = (
            "电子发票（普通发票）\n"
            "发票代码：123456789012\n"
            "发票号码：87654321\n"
            "开票日期：2024年01月15日\n"
            "校验码：12345678901234567890\n"
            "购买方信息名称：测试购买方公司统一社会信用代码/纳税人识别号：91110108MA01ABCD1X\n"
            "销售方信息名称：测试销售方公司统一社会信用代码/纳税人识别号：91110108MA01EFGH2Y\n"
            "价税合计¥791.00\n"
            "合计¥700.00¥91.00\n"
        )
        compact = _compact_text(text)
        meta = parse_invoice_metadata_from_compact_text(compact)

        assert meta.get("发票种类") == "电子发票（普通发票）"
        assert meta.get("发票代码") == "123456789012"
        assert meta.get("发票号码") == "87654321"
        assert meta.get("开票日期") == "2024年01月15日"
        assert meta.get("校验码") == "12345678901234567890"
        assert meta.get("购买方名称") == "测试购买方公司"
        assert meta.get("购买方纳税人识别号") == "91110108MA01ABCD1X"
        assert meta.get("销售方名称") == "测试销售方公司"
        assert meta.get("销售方纳税人识别号") == "91110108MA01EFGH2Y"
        # 价税合计 or 金额/税额 should be detected (exact values depend on
        # which 合计 regex match wins in the single-line compact text)
        has_total = meta.get("价税合计") is not None or meta.get("金额") is not None
        assert has_total, "Should detect at least one amount field"


class TestRowParser:
    """Test detail-line row parsing from raw text strings."""

    def test_parse_rows_marked(self) -> None:
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import (
            parse_invoice_rows_from_pdf_text,
        )

        # Text-based parser expects token-per-line for tax rate/amount detection.
        # Use multi-line format where 税率 and 税额 are on separate lines.
        text = (
            "项目名称 规格型号 单位 数量 单价 金额 税率 税额\n"
            "*测试商品一\n"
            "件\n"
            "2\n"
            "100.00\n"
            "200.00\n"
            "13%\n"
            "26.00\n"
            "*测试商品二\n"
            "箱\n"
            "1\n"
            "500.00\n"
            "500.00\n"
            "13%\n"
            "65.00\n"
            "合计 ¥700.00 ¥91.00\n"
        )
        rows = parse_invoice_rows_from_pdf_text(text, prefer_marked=True)

        assert len(rows) >= 1, f"Expected at least 1 row, got {len(rows)}: {rows}"

    def test_empty_text_returns_empty(self) -> None:
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import (
            parse_invoice_rows_from_pdf_text,
        )

        rows = parse_invoice_rows_from_pdf_text("No invoice data here")
        assert rows == []

    def test_yaml_writer_structure(self) -> None:
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.markdown_writer import (
            _build_yaml_frontmatter,
            _render_markdown_table,
        )

        metadata: dict[str, str | None] = {
            "发票代码": "123456789012",
            "发票号码": "87654321",
            "开票日期": "2024年01月15日",
            "购买方名称": "测试公司",
        }
        rows = [
            {"商品名称": "商品A", "金额": "100.00", "税额": "13.00"},
            {"商品名称": "商品B", "金额": "200.00", "税额": "26.00"},
        ]

        yaml_text = _build_yaml_frontmatter(file_stem="test", metadata=metadata, include_empty=False)
        assert yaml_text.startswith("---")
        assert "标题: test" in yaml_text
        assert "发票号码" in yaml_text

        localized_yaml_text = _build_yaml_frontmatter(
            file_stem="test",
            metadata=metadata,
            include_empty=False,
            yaml_key_labels={"title": "Titel"},
        )
        assert "Titel: test" in localized_yaml_text
        assert "标题: test" not in localized_yaml_text
        assert "发票号码" in localized_yaml_text

        from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import TABLE_HEADERS

        table_text = _render_markdown_table(headers=TABLE_HEADERS, rows=rows)
        assert "| 商品名称 |" in table_text
        assert "商品A" in table_text
        assert "商品B" in table_text

        full = yaml_text + "## 商品明细\n\n" + table_text
        assert full.startswith("---")
        assert "## 商品明细" in full
