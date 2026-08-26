"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

from ._layout_conversions_support import (
    Any,
    Path,
    _assert_yaml_title,
    _build_fake_context,
    _build_runtime_pipeline,
    _load_layout_pdf_old_system_fixture,
    _ocr_success,
    _page_fragments,
    create_image_xps,
    create_minimal_xps,
    docx_semantic_projection,
    png_visual_projection,
    pytest,
    sys,
    tempfile,
    types,
)

pytestmark = pytest.mark.contract


class TestPreprocessChain:
    """Verify that OFD/XPS inputs go through the preprocess layer before
    reaching the downstream converters."""

    def test_ofd_to_md_preprocess_chain(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """OFD→MD: preprocess must be called, then converter gets PDF override."""
        import docwen_plugin_layout.plugin as plugin_module
        from docwen_plugin_layout import LayoutPlugin
        from docwen_plugin_layout.preprocess import PreprocessResult

        ofd_path = tmp_path / "test.ofd"
        ofd_path.write_text("fake ofd content")

        preprocess_calls: list[dict] = []
        converter_calls: list[dict] = []

        def fake_preprocess(input_path, staging_dir, source_format):
            preprocess_calls.append(
                {
                    "input_path": input_path,
                    "source_format": source_format,
                }
            )
            # Simulate successful OFD→PDF conversion
            pdf_path = str(tmp_path / "_preprocess_ofd_fake.pdf")
            Path(pdf_path).write_text("fake pdf content")
            return PreprocessResult(
                effective_input_path=pdf_path,
                original_source_format="ofd",
                effective_source_format="pdf",
                intermediate_artifacts=[pdf_path],
            )

        monkeypatch.setattr(plugin_module, "preprocess_layout_input", fake_preprocess)

        # Also mock the converter so we don't need pymupdf4llm
        def fake_md_convert(self, context, *, input_path_override=None, source_format_override=None):
            converter_calls.append(
                {
                    "input_path_override": input_path_override,
                    "source_format_override": source_format_override,
                }
            )
            from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
            from docwen_core.models.result import ConversionDiagnostic, ConversionMetrics, ConversionResult

            md_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".md")
            Path(md_path).write_text("# Test", encoding="utf-8")
            artifact = ArtifactManifest(
                artifact_id="test-md",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=md_path,
                suggested_name="test.md",
                media_type="text/markdown",
                metadata={"source_format": source_format_override or "pdf"},
                is_primary=True,
            )
            context.workspace.add_artifact(artifact)
            return ConversionResult(
                task_id=context.request.request_id,
                success=True,
                artifacts=[artifact],
                diagnostics=[ConversionDiagnostic(level="info", message="ok", code="OK")],
                metrics=ConversionMetrics(input_bytes=0, output_bytes=0),
            )

        monkeypatch.setattr(
            "docwen_plugin_layout.to_markdown.converter.LayoutToMarkdownConverter.convert",
            fake_md_convert,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(ofd_path),
                staging,
                "md",
                source_format="ofd",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True
            # preprocess must have been called with ofd source
            assert len(preprocess_calls) == 1
            assert preprocess_calls[0]["source_format"] == "ofd"
            # converter must have received PDF override
            assert len(converter_calls) == 1
            assert converter_calls[0]["source_format_override"] == "pdf"
            assert "fake.pdf" in converter_calls[0]["input_path_override"]

    def test_xps_to_png_preprocess_chain(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """XPS→PNG: preprocess must be called, then Image converter gets PDF override."""
        import docwen_plugin_layout.plugin as plugin_module
        from docwen_plugin_layout import LayoutPlugin
        from docwen_plugin_layout.preprocess import PreprocessResult

        xps_path = tmp_path / "test.xps"
        xps_path.write_text("fake xps content")

        preprocess_calls: list[dict] = []

        def fake_preprocess(input_path, staging_dir, source_format):
            preprocess_calls.append(
                {
                    "input_path": input_path,
                    "source_format": source_format,
                }
            )
            pdf_path = str(tmp_path / "_preprocess_xps_fake.pdf")
            Path(pdf_path).write_text("fake pdf content")
            return PreprocessResult(
                effective_input_path=pdf_path,
                original_source_format="xps",
                effective_source_format="pdf",
                intermediate_artifacts=[pdf_path],
            )

        monkeypatch.setattr(plugin_module, "preprocess_layout_input", fake_preprocess)

        # Create a fake PDF that PyMuPDF can open for the image converter
        import fitz

        pdf_path = tmp_path / "fake_input.pdf"
        doc = fitz.open()
        doc.new_page(width=100, height=100)
        doc.save(str(pdf_path))
        doc.close()

        # We need the override path to point to a real PDF
        # The fake_preprocess returns a text file — that won't work with fitz.
        # Instead, make the fake_preprocess point to the real PDF fixture.
        def fake_preprocess_v2(input_path, staging_dir, source_format):
            preprocess_calls.append(
                {
                    "input_path": input_path,
                    "source_format": source_format,
                }
            )
            return PreprocessResult(
                effective_input_path=str(pdf_path),
                original_source_format="xps",
                effective_source_format="pdf",
                intermediate_artifacts=[str(pdf_path)],
            )

        monkeypatch.setattr(plugin_module, "preprocess_layout_input", fake_preprocess_v2)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(xps_path),
                staging,
                "png",
                source_format="xps",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True
            assert len(preprocess_calls) == 1
            assert preprocess_calls[0]["source_format"] == "xps"
            # Output should be PNG
            assert result.artifacts[0].media_type == "image/png"
            assert result.artifacts[0].suggested_name == "test_page_01.png"
            assert "_preprocess" not in result.artifacts[0].suggested_name

    def test_real_xps_to_markdown_uses_original_input_stem(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """XPS preprocessing must not leak its random PDF stem into Markdown."""
        from docwen_plugin_layout import LayoutPlugin

        xps_path = tmp_path / "source-layout.xps"
        create_minimal_xps(xps_path)

        def fake_to_markdown(document: Any, **_kwargs: Any) -> str:
            assert bool(document.is_pdf) is True
            assert bool(document.is_closed) is False
            assert Path(str(document.name)).name.startswith("_preprocess_xps_")
            return "# XPS body\n"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(xps_path),
                staging,
                "md",
                source_format="xps",
                options={"to_md_keep_images": False, "to_md_enable_ocr": False},
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            artifact = next(item for item in result.artifacts if item.media_type == "text/markdown")
            assert artifact.suggested_name == "source-layout.md"
            markdown = Path(artifact.staging_path).read_text(encoding="utf-8")
            _assert_yaml_title(markdown, "title", "source-layout")
            assert "_preprocess_xps_" not in markdown

    def test_real_xps_to_markdown_images_use_original_input_name(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Real XPS image/OCR artifacts must not expose a staging PDF UUID."""
        import shutil

        import docwen_plugin_layout.to_markdown.converter as converter
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        xps_path = tmp_path / "image-pages.xps"
        expected = _load_layout_pdf_old_system_fixture()["real_xps_to_markdown_probe"]["projects"]["docwen-current"]
        create_image_xps(xps_path)
        monkeypatch.setattr(
            converter,
            "run_ocr_outcome",
            lambda image_path, **_kwargs: _ocr_success(f"OCR::{Path(image_path).stem}"),
        )

        task_manager, ws_manager, ws_root = _build_runtime_pipeline()
        try:
            output_dir = tmp_path / "output_xps_markdown"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id="layout-real-xps-markdown",
                input_refs=[
                    FileRef(
                        path=str(xps_path),
                        format="xps",
                        category="layout",
                        size_bytes=xps_path.stat().st_size,
                    )
                ],
                target_format="md",
                output_policy=OutputPolicy(output_dir=str(output_dir)),
                options={
                    "to_md_keep_images": True,
                    "to_md_enable_ocr": True,
                    "image_mode": "file",
                },
            )
            result = task_manager.execute_single(request)
        finally:
            ws_manager.cleanup_all()
            shutil.rmtree(ws_root, ignore_errors=True)

        assert result.success is True, f"unexpected error: {result.error}"
        page_fragments = _page_fragments(result)
        image_artifacts = [artifact for artifact in result.artifacts if artifact.kind == "image"]
        deliverables = [artifact for artifact in result.artifacts if artifact.kind != "manifest"]
        assert [
            artifact.metadata.get("source_suggested_name", artifact.suggested_name) for artifact in deliverables
        ] == expected["artifact_suggested_names"]
        assert [
            {
                "page_index": artifact.metadata["page_index"],
                "page_count": artifact.metadata["page_count"],
                "source_page": artifact.metadata["source_page"],
                "ocr_status": artifact.metadata["ocr_status"],
            }
            for artifact in page_fragments
        ] == expected["page_fragment_metadata"]
        assert [artifact.metadata["source_page"] for artifact in image_artifacts] == [1, 2]
        markdown_path = Path(result.artifacts[0].staging_path)
        markdown = markdown_path.read_text(encoding="utf-8")
        assert "![[image-pages.xps-0001-00.png]]" in markdown
        assert "![[image-pages.xps-0002-00.png]]" in markdown
        assert "OCR::image-pages.xps-0001-00" not in markdown
        assert "OCR::image-pages.xps-0002-00" not in markdown
        assert ("OCR::" in markdown) is expected["primary_contains_ocr"]
        assert [Path(artifact.staging_path).read_text(encoding="utf-8") for artifact in page_fragments] == [
            "OCR::_ocr_page_1\n",
            "OCR::_ocr_page_2\n",
        ]
        assert "_preprocess_xps_" not in markdown
        assert all("_preprocess_xps_" not in artifact.suggested_name for artifact in result.artifacts)
        image_paths = [
            Path(artifact.staging_path) for artifact in result.artifacts if artifact.media_type == "image/png"
        ]
        assert png_visual_projection(image_paths) == expected["observed_same_environment_images"]

    def test_real_xps_to_docx_pdf2docx_fallback_matches_old_system_projection(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Real XPS→DOCX fallback preserves the normalized old-system artifact semantics."""
        import shutil

        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_core.office_bridge import BridgeResult

        xps_path = tmp_path / "image-pages.xps"
        create_image_xps(xps_path)
        probe = _load_layout_pdf_old_system_fixture()["real_xps_to_docx_probe"]
        expected = probe["expected_semantic_projection"]
        current = probe["projects"]["docwen-current"]

        monkeypatch.setattr(
            converter,
            "_convert_pdf_with_configured_office_priority",
            lambda *_args, **_kwargs: BridgeResult(
                False,
                message="controlled fallback after external Word timeout",
            ),
        )

        task_manager, ws_manager, ws_root = _build_runtime_pipeline()
        try:
            output_dir = tmp_path / "output_xps_docx"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id="layout-real-xps-docx",
                input_refs=[
                    FileRef(
                        path=str(xps_path),
                        format="xps",
                        category="layout",
                        size_bytes=xps_path.stat().st_size,
                    )
                ],
                target_format="docx",
                output_policy=OutputPolicy(output_dir=str(output_dir)),
            )
            result = task_manager.execute_single(request)
        finally:
            ws_manager.cleanup_all()
            shutil.rmtree(ws_root, ignore_errors=True)

        assert result.success is True, f"unexpected error: {result.error}"
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        assert artifact.suggested_name == current["artifact_suggested_name"]
        assert "_preprocess" not in artifact.suggested_name
        assert result.metrics.extra["engine"] == "pdf2docx"
        assert result.metrics.extra["backend"] == "pdf2docx"
        assert result.metrics.extra["output_dir"] == str(output_dir)
        assert Path(artifact.staging_path).parent == output_dir
        assert Path(artifact.staging_path).is_file()
        assert docx_semantic_projection(artifact.staging_path) == expected

    def test_ofd_to_docx_preprocess_chain(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """OFD→DOCX: preprocess must be called, then Document converter gets PDF override."""
        import docwen_plugin_layout.plugin as plugin_module
        from docwen_plugin_layout import LayoutPlugin
        from docwen_plugin_layout.preprocess import PreprocessResult

        ofd_path = tmp_path / "test.ofd"
        ofd_path.write_text("fake ofd content")

        preprocess_calls: list[dict] = []

        def fake_preprocess(input_path, staging_dir, source_format):
            preprocess_calls.append(
                {
                    "input_path": input_path,
                    "source_format": source_format,
                }
            )
            pdf_path = str(tmp_path / "_preprocess_ofd_fake.pdf")
            Path(pdf_path).write_text("fake pdf content")
            return PreprocessResult(
                effective_input_path=pdf_path,
                original_source_format="ofd",
                effective_source_format="pdf",
                intermediate_artifacts=[pdf_path],
            )

        monkeypatch.setattr(plugin_module, "preprocess_layout_input", fake_preprocess)

        # Mock the final PDF->DOCX fallback. This test focuses on the OFD
        # preprocess chain, so external PDF import is disabled via config.
        import docwen_plugin_layout.to_document.converter as doc_conv
        from docwen_core.office_bridge import BridgeResult

        def fake_pdf2docx(
            input_path: Path,
            output_path: Path,
            *,
            cancellation: Any = None,
        ) -> BridgeResult:
            assert cancellation is not None
            Path(output_path).write_bytes(b"docx-by-pdf2docx")
            return BridgeResult(success=True, output_path=str(output_path), backend="pdf2docx")

        monkeypatch.setattr(doc_conv, "_convert_pdf_with_pdf2docx", fake_pdf2docx)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(ofd_path),
                staging,
                "docx",
                source_format="ofd",
                config_values={"software": {"special_conversions": {"pdf_to_office": []}}},
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True
            assert len(preprocess_calls) == 1
            assert preprocess_calls[0]["source_format"] == "ofd"
            assert (
                result.artifacts[0].media_type
                == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            assert result.artifacts[0].suggested_name == "test.docx"
            assert "_preprocess" not in result.artifacts[0].suggested_name
            assert result.metrics.extra["engine"] == "pdf2docx"

    def test_pdf_to_md_no_preprocess_needed(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """PDF→MD: preprocess is called but passes through (effective_source stays pdf)."""
        import fitz

        import docwen_plugin_layout.plugin as plugin_module
        from docwen_plugin_layout import LayoutPlugin

        pdf_path = tmp_path / "test.pdf"
        doc = fitz.open()
        doc.new_page(width=100, height=100)
        doc.save(str(pdf_path))
        doc.close()

        # Track that preprocess was called (passthrough)
        original_preprocess = plugin_module.preprocess_layout_input
        preprocess_calls: list[dict] = []

        def tracking_preprocess(input_path, staging_dir, source_format):
            preprocess_calls.append({"source_format": source_format})
            return original_preprocess(input_path, staging_dir, source_format)

        monkeypatch.setattr(plugin_module, "preprocess_layout_input", tracking_preprocess)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                source_format="pdf",
                options={"to_md_keep_images": False},
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(preprocess_calls) == 1
            assert preprocess_calls[0]["source_format"] == "pdf"

    def test_pdf_to_md_real_text_smoke(self, sample_pdf_path: Path) -> None:
        """Real pymupdf4llm smoke: PDF text should become a Markdown artifact."""
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "md",
                source_format="pdf",
                options={"to_md_keep_images": False, "to_md_enable_ocr": False},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "Test document" in md_text
            assert md_artifact.suggested_name == "sample.md"
            assert result.metrics.extra["image_count"] == 0
            assert result.metrics.extra["ocr_enabled"] is False
