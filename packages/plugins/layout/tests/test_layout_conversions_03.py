"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

from ._layout_conversions_support import (
    Any,
    Path,
    _assert_open_pdf_document,
    _assert_yaml_title,
    _build_fake_context,
    _build_runtime_pipeline,
    _create_metadata_pdf,
    _create_text_pdf,
    _document_node_root,
    _load_layout_pdf_old_system_fixture,
    _pdf_metadata_projection,
    _pdf_page_texts,
    create_minimal_xps,
    pdf_visual_projection,
    png_visual_projection,
    pytest,
    raster_visual_projection,
    sys,
    tempfile,
    types,
)

pytestmark = pytest.mark.contract


class TestOldSystemLayoutPdfFixture:
    def test_runtime_pdf_passthrough_uses_admitted_format_despite_ofd_suffix(self, tmp_path: Path) -> None:
        """Runtime dispatch and converter parsing share the admitted PDF identity."""
        import shutil

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        misleading_path = tmp_path / "admitted-pdf.ofd"
        _create_text_pdf(misleading_path, ["ADMITTED-PDF-CONTENT"])
        output_dir = tmp_path / "out"
        output_dir.mkdir()
        task_manager, ws_manager, ws_root = _build_runtime_pipeline()
        try:
            request = ConversionRequest(
                request_id="layout-admitted-pdf-wrong-suffix",
                input_refs=[
                    FileRef(
                        path=str(misleading_path),
                        format="pdf",
                        category="layout",
                        size_bytes=misleading_path.stat().st_size,
                    )
                ],
                target_format="pdf",
                output_policy=OutputPolicy(output_dir=str(output_dir)),
            )
            result = task_manager.execute_single(request)
        finally:
            ws_manager.cleanup_all()
            shutil.rmtree(ws_root, ignore_errors=True)

        assert result.success is True, f"unexpected error: {result.error}"
        artifact = result.artifacts[0]
        assert artifact.metadata["source_format"] == "pdf"
        assert _pdf_page_texts(artifact.staging_path) == ["ADMITTED-PDF-CONTENT"]

    def test_pdf_to_md_matches_old_system_semantic_fixture(self, tmp_path: Path) -> None:
        """PDF→MD should preserve the old-system focused text extraction semantics."""
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        fixture = _load_layout_pdf_old_system_fixture()
        pdf_input = fixture["input_pdf"]
        pdf_path = tmp_path / pdf_input["name"]
        _create_text_pdf(pdf_path, pdf_input["page_texts"])
        expected = fixture["expected"]["pdf_to_md"]
        current = fixture["projects"]["docwen-current"]["pdf_to_md"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                source_format="pdf",
                options={"to_md_keep_images": False, "to_md_enable_ocr": False},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is expected["success"], f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            _assert_yaml_title(md_text, "title", pdf_path.stem)
            for text in expected["markdown_contains"]:
                assert text in md_text
            assert md_artifact.suggested_name == current["suggested_name"]
            assert md_artifact.metadata == current["metadata"]
            for key, value in current["metrics"].items():
                assert result.metrics.extra[key] == value

    def test_pdf_to_md_passes_live_document_and_closes_it_for_unknown_suffix(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """pymupdf4llm receives one explicit PDF document, never the user path."""
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        misleading_path = tmp_path / "actual-pdf-content.unknown"
        _create_text_pdf(misleading_path, ["CONTENT FIRST PDF TO MARKDOWN"])
        observed_documents: list[Any] = []

        def fake_to_markdown(document: Any, **_kwargs: Any) -> str:
            _assert_open_pdf_document(document, misleading_path)
            observed_documents.append(document)
            return "CONTENT FIRST PDF TO MARKDOWN"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))

        staging = tmp_path / "markdown-staging"
        staging.mkdir()
        context = _build_fake_context(
            str(misleading_path),
            str(staging),
            "md",
            source_format="pdf",
            options={"to_md_keep_images": False, "to_md_enable_ocr": False},
        )
        result = LayoutToMarkdownConverter().convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        assert "CONTENT FIRST PDF TO MARKDOWN" in Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert len(observed_documents) == 1
        assert bool(observed_documents[0].is_closed) is True

    def test_pdf_to_md_yaml_frontmatter_consumes_locale_title_label(self, tmp_path: Path) -> None:
        """PDF→MD YAML frontmatter should consume app-resolved locale labels."""
        import shutil

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        pdf_path = tmp_path / "layout_locale_probe.pdf"
        _create_text_pdf(pdf_path, ["LAYOUT-LOCALE-PROBE"])
        task_manager, ws_manager, ws_root = _build_runtime_pipeline()
        try:
            output_dir = tmp_path / "output_layout_locale_yaml"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id="layout-pdf-locale-finalizer",
                input_refs=[
                    FileRef(
                        path=str(pdf_path),
                        format="pdf",
                        category="layout",
                        size_bytes=pdf_path.stat().st_size,
                    )
                ],
                target_format="md",
                output_policy=OutputPolicy(output_dir=str(output_dir)),
                options={
                    "to_md_keep_images": False,
                    "to_md_enable_ocr": False,
                    "yaml_key_labels": {"title": "Titel"},
                },
            )

            result = task_manager.execute_single(request)
        finally:
            ws_manager.cleanup_all()
            shutil.rmtree(ws_root, ignore_errors=True)

        assert result.success is True, f"unexpected error: {result.error}"
        md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
        md_path = Path(md_artifact.staging_path)
        node_root = _document_node_root(md_path, output_dir)
        assert md_path.name == f"{node_root.name}.md"
        assert md_path.exists()
        md_text = md_path.read_text(encoding="utf-8")
        assert str(Path(ws_root)) not in md_text
        _assert_yaml_title(md_text, "Titel", "layout_locale_probe")
        assert "LAYOUT-LOCALE-PROBE" in md_text

    def test_layout_pdf_passthrough_matches_old_system_semantic_fixture(self, tmp_path: Path) -> None:
        """PDF→PDF should preserve PDF content while using current runtime artifacts."""
        from docwen_plugin_layout.to_pdf.converter import LayoutToPdfConverter

        fixture = _load_layout_pdf_old_system_fixture()
        pdf_input = fixture["input_pdf"]
        pdf_path = tmp_path / pdf_input["name"]
        _create_text_pdf(pdf_path, pdf_input["page_texts"])
        expected = fixture["expected"]["layout_to_pdf"]
        current = fixture["projects"]["docwen-current"]["layout_to_pdf"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(pdf_path), staging, "pdf", source_format="pdf")
            result = LayoutToPdfConverter().convert(context)

            assert result.success is expected["success"], f"unexpected error: {result.error}"
            artifact = result.artifacts[0]
            assert artifact.media_type == current["artifact_media_type"]
            assert artifact.suggested_name == current["suggested_name"]
            assert artifact.metadata == current["metadata"]
            assert result.metrics.extra == current["metrics"]
            assert Path(artifact.staging_path).read_bytes().startswith(expected["pdf_magic"].encode("ascii"))
            assert _pdf_page_texts(artifact.staging_path) == expected["page_texts"]

    def test_layout_pdf_metadata_passthrough_matches_old_system_projection(self, tmp_path: Path) -> None:
        """PDF→PDF passthrough/copy should preserve selected PDF metadata fields."""
        from docwen_plugin_layout.to_pdf.converter import LayoutToPdfConverter

        fixture = _load_layout_pdf_old_system_fixture()
        probe = fixture["metadata_passthrough_probe"]
        pdf_input = probe["input_pdf"]
        pdf_path = tmp_path / pdf_input["name"]
        _create_metadata_pdf(pdf_path, pdf_input)
        expected = probe["projects"]["docwen-current"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(pdf_path), staging, "pdf", source_format="pdf")
            result = LayoutToPdfConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.media_type == expected["artifact_media_type"]
            assert artifact.suggested_name == expected["suggested_name"]
            assert artifact.metadata == expected["artifact_metadata"]
            assert result.metrics.extra == expected["metrics"]
            assert _pdf_metadata_projection(artifact.staging_path) == {
                "pdf_magic": expected["pdf_magic"],
                "page_count": expected["page_count"],
                "page_texts": expected["page_texts"],
                "metadata": expected["metadata"],
            }

    def test_real_xps_to_pdf_matches_old_system_projection(self, tmp_path: Path) -> None:
        """A real two-page XPS should retain old-system PDF visual semantics."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_layout_pdf_old_system_fixture()
        probe = fixture["real_xps_to_pdf_probe"]
        expected = probe["projects"]["docwen-current"]
        xps_path = tmp_path / probe["input_xps"]["name"]
        create_minimal_xps(xps_path)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xps_path), staging, "pdf", source_format="xps")
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.media_type == expected["artifact_media_type"]
            assert artifact.suggested_name == expected["suggested_name"]
            assert artifact.metadata == expected["artifact_metadata"]
            assert result.metrics.extra == expected["metrics"]
            assert pdf_visual_projection(artifact.staging_path) == probe["expected_projection"]

    def test_real_xps_to_png_matches_old_system_projection(self, tmp_path: Path) -> None:
        """Real XPS→PNG should match old pixels and finalize source-owned names."""
        import shutil

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        fixture = _load_layout_pdf_old_system_fixture()
        probe = fixture["real_xps_to_png_probe"]
        expected = probe["projects"]["docwen-current"]
        xps_path = tmp_path / probe["input_xps"]["name"]
        create_minimal_xps(xps_path)

        task_manager, ws_manager, ws_root = _build_runtime_pipeline()
        try:
            output_dir = tmp_path / "output_xps_png"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id="layout-real-xps-png",
                input_refs=[
                    FileRef(
                        path=str(xps_path),
                        format="xps",
                        category="layout",
                        size_bytes=xps_path.stat().st_size,
                    )
                ],
                target_format="png",
                output_policy=OutputPolicy(output_dir=str(output_dir)),
                options={"render_dpi": probe["input_xps"]["render_dpi"]},
            )
            result = task_manager.execute_single(request)
        finally:
            ws_manager.cleanup_all()
            shutil.rmtree(ws_root, ignore_errors=True)

        assert result.success is True, f"unexpected error: {result.error}"
        deliverables = [artifact for artifact in result.artifacts if artifact.kind != "manifest"]
        assert [
            artifact.metadata.get("source_suggested_name", artifact.suggested_name) for artifact in deliverables
        ] == expected["artifact_suggested_names"]
        assert all(artifact.media_type == expected["artifact_media_type"] for artifact in result.artifacts)
        assert [artifact.metadata for artifact in result.artifacts] == expected["artifact_metadata"]
        output_paths = [Path(artifact.staging_path) for artifact in result.artifacts]
        assert all(path.parent == output_dir and path.exists() for path in output_paths)
        assert png_visual_projection(output_paths) == probe["expected_projection"]["images"]
        assert any(diagnostic.code == "FINALIZER_DONE" for diagnostic in result.diagnostics)

    @pytest.mark.parametrize("target_format", ["jpg", "tif"])
    def test_real_xps_to_jpg_tif_matches_old_system_projection(
        self,
        tmp_path: Path,
        target_format: str,
    ) -> None:
        """Real XPS→JPG/TIF should retain old pixels and current output ownership."""
        import shutil

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        fixture = _load_layout_pdf_old_system_fixture()
        probe = fixture["real_xps_to_jpg_tif_probe"]
        expected = probe["projects"]["docwen-current"][target_format]
        xps_path = tmp_path / probe["input_xps"]["name"]
        create_minimal_xps(xps_path)

        task_manager, ws_manager, ws_root = _build_runtime_pipeline()
        try:
            output_dir = tmp_path / f"output_xps_{target_format}"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id=f"layout-real-xps-{target_format}",
                input_refs=[
                    FileRef(
                        path=str(xps_path),
                        format="xps",
                        category="layout",
                        size_bytes=xps_path.stat().st_size,
                    )
                ],
                target_format=target_format,
                output_policy=OutputPolicy(output_dir=str(output_dir)),
                options={"render_dpi": probe["input_xps"]["render_dpi"]},
            )
            result = task_manager.execute_single(request)
        finally:
            ws_manager.cleanup_all()
            shutil.rmtree(ws_root, ignore_errors=True)

        assert result.success is True, f"unexpected error: {result.error}"
        assert [artifact.suggested_name for artifact in result.artifacts] == expected["artifact_suggested_names"]
        assert [artifact.media_type for artifact in result.artifacts] == expected["artifact_media_types"]
        assert [artifact.metadata for artifact in result.artifacts] == expected["artifact_metadata"]
        output_paths = [Path(artifact.staging_path) for artifact in result.artifacts]
        assert all(path.parent == output_dir and path.exists() for path in output_paths)
        assert raster_visual_projection(output_paths) == probe["expected_projection"][target_format]
        assert any(diagnostic.code == "FINALIZER_DONE" for diagnostic in result.diagnostics)

    def test_layout_pdf_passthrough_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path) -> None:
        """PDF→PDF passthrough should finalize the copied PDF into the user output dir."""
        import shutil

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        fixture = _load_layout_pdf_old_system_fixture()
        pdf_input = fixture["input_pdf"]
        pdf_path = tmp_path / pdf_input["name"]
        _create_text_pdf(pdf_path, pdf_input["page_texts"])
        expected = fixture["expected"]["layout_to_pdf"]
        current = fixture["projects"]["docwen-current"]["layout_to_pdf"]

        task_manager, ws_manager, ws_root = _build_runtime_pipeline()
        try:
            output_dir = tmp_path / "output_layout_pdf"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id="layout-pdf-passthrough-finalizer-old-system-fixture",
                input_refs=[
                    FileRef(
                        path=str(pdf_path),
                        format="pdf",
                        category="layout",
                        size_bytes=pdf_path.stat().st_size,
                    )
                ],
                target_format="pdf",
                output_policy=OutputPolicy(output_dir=str(output_dir)),
                options={},
            )

            result = task_manager.execute_single(request)
        finally:
            ws_manager.cleanup_all()
            shutil.rmtree(ws_root, ignore_errors=True)

        assert result.success, f"unexpected error: {result.error}"
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        output_path = Path(artifact.staging_path)
        assert output_path.parent == output_dir
        assert output_path.name == current["suggested_name"]
        assert artifact.media_type == current["artifact_media_type"]
        assert artifact.metadata == current["metadata"]
        assert output_path.read_bytes().startswith(expected["pdf_magic"].encode("ascii"))
        assert _pdf_page_texts(output_path) == expected["page_texts"]
        assert any(d.code == "PDF-CONVERT-OK" for d in result.diagnostics)
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
        assert str(ws_root) not in str(output_path)
