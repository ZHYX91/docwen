"""Focused tests split from test_invoice_conversions.py."""

from __future__ import annotations

from ._invoice_conversions_support import (
    Path,
    _build_fake_context,
    _convert_invoice_fixture,
    _execute_invoice_runtime_fixture,
    _load_invoice_cn_old_system_fixture,
    _parse_invoice_markdown,
    os,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.contract


class TestPdfInvoiceToMarkdown:
    def test_runtime_pdf_invoice_uses_admitted_format_despite_ofd_suffix(
        self, sample_invoice_pdf_path: Path, tmp_path: Path
    ) -> None:
        """A post-admission PDF invoice must not be routed by its OFD suffix."""
        misleading_path = tmp_path / "admitted-pdf-invoice.ofd"
        misleading_path.write_bytes(sample_invoice_pdf_path.read_bytes())
        output_dir = tmp_path / "out"
        output_dir.mkdir()

        result = _execute_invoice_runtime_fixture(
            input_path=misleading_path,
            source_format="pdf",
            output_dir=output_dir,
            workspace_root=tmp_path / "workspace",
            request_id="invoice-admitted-pdf-wrong-suffix",
        )

        assert result.success is True, f"unexpected error: {result.error}"
        assert result.artifacts[0].metadata["source_format"] == "pdf"

    def test_invoice_cn_matches_old_system_semantic_fixture(
        self,
        sample_invoice_pdf_path: Path,
        sample_invoice_ofd_path: Path,
    ) -> None:
        """Current invoice_cn should preserve old-system PDF/OFD core semantics."""
        fixture = _load_invoice_cn_old_system_fixture()
        assert fixture["golden_id"] == "GOLDEN-023"

        cases = {
            "pdf": (sample_invoice_pdf_path, "pdf"),
            "ofd": (sample_invoice_ofd_path, "ofd"),
        }
        for case_name, (path, source_format) in cases.items():
            result, markdown = _convert_invoice_fixture(path, source_format)
            yaml_fields, table_rows = _parse_invoice_markdown(markdown)
            expected = fixture["expected_semantics"][case_name]
            artifact = result.artifacts[0]

            assert artifact.media_type == "text/markdown"
            assert artifact.metadata["source_format"] == source_format
            assert artifact.metadata["row_count"] == expected["row_count"]
            assert artifact.metadata["yaml_fields"] == fixture["schema_field_count"]
            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)

            for key, value in expected["yaml_subset"].items():
                assert yaml_fields[key] == value
            for token in expected["required_markdown_tokens"]:
                assert token in markdown
            if case_name == "ofd":
                assert table_rows == expected["table_row_lines"]

        for project_name in ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"):
            project = fixture["projects"][project_name]
            assert project["success"] is True
            assert project["schema_field_count"] == fixture["schema_field_count"]
            assert project["pdf"]["matches_expected"] is True
            assert project["ofd"]["matches_expected"] is True

    def test_invoice_cn_pdf_fixture_finalizes_through_runtime(
        self,
        sample_invoice_pdf_path: Path,
        tmp_path: Path,
    ) -> None:
        """invoice_cn PDF output should be finalized into the user output dir."""
        fixture = _load_invoice_cn_old_system_fixture()
        expected = fixture["expected_semantics"]["pdf"]
        output_dir = tmp_path / "out"
        output_dir.mkdir()
        workspace_root = tmp_path / "workspace"

        result = _execute_invoice_runtime_fixture(
            input_path=sample_invoice_pdf_path,
            source_format="pdf",
            output_dir=output_dir,
            workspace_root=workspace_root,
            request_id="invoice-cn-pdf-finalizer-old-system-fixture",
            options={"yaml_key_labels": {"title": "Titel"}},
        )

        assert result.success, f"unexpected error: {result.error}"
        assert len(result.artifacts) == 2
        artifact = next(item for item in result.artifacts if item.media_type == "text/markdown")
        manifest = next(
            item for item in result.artifacts if item.media_type == "application/vnd.docwen.document-node+json"
        )
        md_path = Path(artifact.staging_path)
        node_root = md_path.parent
        assert node_root.parent == output_dir
        assert node_root.name.startswith("invoice_")
        assert node_root.name.endswith("_fromPdf")
        assert md_path.name == f"{node_root.name}.md"
        assert artifact.logical_path == f"{node_root.name}/{node_root.name}.md"
        assert Path(manifest.staging_path) == node_root / "docwen-node.json"
        assert manifest.logical_path == f"{node_root.name}/docwen-node.json"
        assert artifact.media_type == "text/markdown"
        assert {key: artifact.metadata[key] for key in ("source_format", "row_count", "yaml_fields")} == {
            "source_format": "pdf",
            "row_count": expected["row_count"],
            "yaml_fields": fixture["schema_field_count"],
        }
        assert artifact.metadata["document_node_schema"] == "docwen.document_node.v1"
        assert artifact.metadata["document_node_role"] == "primary"
        assert artifact.metadata["document_node_committed"] is True
        assert any(d.code == "INVOICE-OK" for d in result.diagnostics)
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
        markdown = md_path.read_text(encoding="utf-8")
        yaml_fields, table_rows = _parse_invoice_markdown(markdown)
        assert yaml_fields["Titel"] == "invoice"
        assert "标题" not in yaml_fields
        for key, value in expected["yaml_subset"].items():
            assert yaml_fields[key] == value
        for token in expected["required_markdown_tokens"]:
            assert token in markdown
        assert table_rows
        assert str(workspace_root) not in markdown

    def test_invoice_cn_ofd_fixture_finalizes_through_runtime(
        self,
        sample_invoice_ofd_path: Path,
        tmp_path: Path,
    ) -> None:
        """invoice_cn OFD output should be finalized into the user output dir."""
        fixture = _load_invoice_cn_old_system_fixture()
        expected = fixture["expected_semantics"]["ofd"]
        output_dir = tmp_path / "out"
        output_dir.mkdir()
        workspace_root = tmp_path / "workspace"

        result = _execute_invoice_runtime_fixture(
            input_path=sample_invoice_ofd_path,
            source_format="ofd",
            output_dir=output_dir,
            workspace_root=workspace_root,
            request_id="invoice-cn-ofd-finalizer-old-system-fixture",
        )

        assert result.success, f"unexpected error: {result.error}"
        assert len(result.artifacts) == 2
        artifact = next(item for item in result.artifacts if item.media_type == "text/markdown")
        manifest = next(
            item for item in result.artifacts if item.media_type == "application/vnd.docwen.document-node+json"
        )
        md_path = Path(artifact.staging_path)
        node_root = md_path.parent
        assert node_root.parent == output_dir
        assert node_root.name.startswith("invoice_")
        assert node_root.name.endswith("_fromOfd")
        assert md_path.name == f"{node_root.name}.md"
        assert artifact.logical_path == f"{node_root.name}/{node_root.name}.md"
        assert Path(manifest.staging_path) == node_root / "docwen-node.json"
        assert manifest.logical_path == f"{node_root.name}/docwen-node.json"
        assert artifact.media_type == "text/markdown"
        assert {key: artifact.metadata[key] for key in ("source_format", "row_count", "yaml_fields")} == {
            "source_format": "ofd",
            "row_count": expected["row_count"],
            "yaml_fields": fixture["schema_field_count"],
        }
        assert artifact.metadata["document_node_schema"] == "docwen.document_node.v1"
        assert artifact.metadata["document_node_role"] == "primary"
        assert artifact.metadata["document_node_committed"] is True
        assert any(d.code == "INVOICE-OK" for d in result.diagnostics)
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
        markdown = md_path.read_text(encoding="utf-8")
        yaml_fields, table_rows = _parse_invoice_markdown(markdown)
        for key, value in expected["yaml_subset"].items():
            assert yaml_fields[key] == value
        assert table_rows == expected["table_row_lines"]
        assert str(workspace_root) not in markdown

    def test_basic_pdf_invoice_conversion(self, sample_invoice_pdf_path: Path) -> None:
        """A PDF invoice → MD should produce YAML frontmatter + table."""
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
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.media_type == "text/markdown"
            assert artifact.suggested_name == "invoice.md"
            assert artifact.is_primary is True
            assert os.path.isfile(artifact.staging_path)

            content = Path(artifact.staging_path).read_text(encoding="utf-8")
            assert len(content) > 0, "Markdown output should not be empty"

            # YAML frontmatter structure
            assert content.startswith("---"), "Output should start with YAML frontmatter"
            assert "---" in content[3:], "YAML frontmatter should have closing delimiter"

            # Table structure (always present even if no rows detected)
            assert "## 商品明细" in content, "Should contain 商品明细 section header"
            assert "| 商品名称 |" in content, "Should contain table header"
            assert "| --- |" in content, "Should contain table separator"

            # Metrics
            assert result.metrics.input_bytes > 0
            assert result.metrics.output_bytes > 0
            assert result.metrics.extra.get("source_format") == "pdf"
            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)

    def test_pdf_invoice_yaml_contains_schema_fields(self, sample_invoice_pdf_path: Path) -> None:
        """The YAML frontmatter should include all 20 invoice schema fields."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import (
            INVOICE_CN_YAML_SCHEMA,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_invoice_pdf_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="pdf",
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True
            content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")

            for field in INVOICE_CN_YAML_SCHEMA:
                assert field in content, f"YAML frontmatter should contain field: {field}"

    def test_pdf_invoice_rows_have_expected_structure(self, sample_invoice_pdf_path: Path) -> None:
        """The table in the output should have 8 columns with proper header."""
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
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True
            content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")

            # The table header should have 8 columns
            table_header_line = None
            for line in content.splitlines():
                if "商品名称" in line and line.startswith("|"):
                    table_header_line = line
                    break

            assert table_header_line is not None, "Should have a table header"
            cols = [c.strip() for c in table_header_line.split("|") if c.strip()]
            assert len(cols) == 8, f"Expected 8 table columns, got {len(cols)}: {cols}"

    def test_pdf_plugin_dispatch_to_converter(self, sample_invoice_pdf_path: Path) -> None:
        """The InvoicePlugin.dispatch should route pdf→md+invoice_cn to converter."""
        from docwen_plugin_optimizer_invoice_cn import InvoicePlugin

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_invoice_pdf_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="pdf",
            )
            result = InvoicePlugin().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)

    def test_cancellation_before_execution(self, sample_invoice_pdf_path: Path) -> None:
        """A pre-cancelled context should raise before any work is done."""
        from docwen_core.errors import CancellationRequested
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
                pre_cancelled=True,
            )
            with pytest.raises(CancellationRequested):
                InvoiceCnConverter().convert(context)


class TestOfdInvoiceToMarkdown:
    def test_ofd_invoice_uses_admitted_format_despite_pdf_suffix(
        self, sample_invoice_ofd_path: Path, tmp_path: Path
    ) -> None:
        """The OFD parser is selected from FileRef.format, not ``.pdf``."""
        misleading_path = tmp_path / "admitted-ofd-invoice.pdf"
        misleading_path.write_bytes(sample_invoice_ofd_path.read_bytes())

        result, markdown = _convert_invoice_fixture(misleading_path, "ofd")

        assert result.artifacts[0].metadata["source_format"] == "ofd"
        assert markdown.startswith("---\n")

    def test_ofd_invoice_with_xml(self, sample_invoice_ofd_path: Path) -> None:
        """OFD with InvoiceData.xml → MD should produce structured output."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_invoice_ofd_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="ofd",
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.media_type == "text/markdown"
            assert artifact.is_primary is True

            content = Path(artifact.staging_path).read_text(encoding="utf-8")

            # OFD InvoiceData.xml should give us these fields
            assert "发票代码" in content
            assert "发票号码" in content
            assert "87654321" in content  # invoice number from XML
            assert "测试购买方公司" in content
            assert "测试销售方公司" in content
            assert "700.00" in content  # total amount
            assert "91.00" in content  # total tax
            assert "791.00" in content  # amount with tax

            # Table rows
            assert "测试商品一" in content
            assert "测试商品二" in content

            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)

    def test_ofd_invoice_fallback_content_xml(self, sample_invoice_ofd_fallback_path: Path) -> None:
        """OFD without InvoiceData.xml → MD using content.xml fallback."""
        from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
            InvoiceCnConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_invoice_ofd_fallback_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="ofd",
            )
            result = InvoiceCnConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 1

            content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")

            # Should have YAML frontmatter and table
            assert content.startswith("---")
            assert "## 商品明细" in content

            # Some fields should be detected from content.xml text
            assert "87654321" in content  # invoice number

    def test_ofd_plugin_dispatch(self, sample_invoice_ofd_path: Path) -> None:
        """The InvoicePlugin.dispatch should route ofd→md+invoice_cn to converter."""
        from docwen_plugin_optimizer_invoice_cn import InvoicePlugin

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_invoice_ofd_path),
                staging,
                "md",
                action_name="invoice_cn",
                source_format="ofd",
            )
            result = InvoicePlugin().convert(context)

            assert result.success is True
            assert any(d.code == "INVOICE-OK" for d in result.diagnostics)
