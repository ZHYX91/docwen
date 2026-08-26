"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

from ._layout_conversions_support import (
    Any,
    Path,
    _assert_open_pdf_document,
    _build_fake_context,
    _create_text_pdf,
    _write_valid_png,
    create_minimal_xps,
    pytest,
    sys,
    tempfile,
    types,
)

pytestmark = pytest.mark.contract


class TestPreprocessChain:
    """Verify that OFD/XPS inputs go through the preprocess layer before
    reaching the downstream converters."""

    def test_pdf_to_md_base64_image_mode_inlines_images(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """base64 image_mode should inline extracted images and avoid image artifacts."""
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "page-1.png"))
            return "![[sample_images/page-1.png]]"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": True, "image_mode": "base64", "image_link_style": "markdown_embed"},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "data:image/png;base64," in md_text
            assert "page-1.png" not in md_text
            assert not any(artifact.kind == "image" for artifact in result.artifacts)

    def test_pdf_to_md_image_mode_comes_from_request_snapshot(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """PDF→MD projects image mode from the admitted request snapshot."""
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "page-1.png"))
            return "![[sample_images/page-1.png]]"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        staging = tmp_path / "layout-request-snapshot-staging"
        staging.mkdir()
        context = _build_fake_context(
            str(pdf_path),
            str(staging),
            "md",
            options={"to_md_keep_images": True, "image_link_style": "markdown_embed"},
            config_snapshot={"export": {"to_md_image_extraction_mode": "base64"}},
        )
        result = LayoutToMarkdownConverter().convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
        md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
        assert "data:image/png;base64," in md_text
        assert not any(artifact.kind == "image" for artifact in result.artifacts)

    def test_pdf_to_md_nonempty_snapshot_is_authoritative(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "snapshot.pdf"
        _create_text_pdf(pdf_path, ["Text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "page-1.png"))
            return "![[snapshot_images/page-1.png]]"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        staging = tmp_path / "layout-request-snapshot-staging"
        staging.mkdir()
        context = _build_fake_context(
            str(pdf_path),
            str(staging),
            "md",
            options={"to_md_keep_images": True, "image_link_style": "markdown_embed"},
            config_snapshot={"export": {"to_md_image_extraction_mode": "base64"}},
        )
        result = LayoutToMarkdownConverter().convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
        md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
        assert "data:image/png;base64," in md_text
        assert not any(artifact.kind == "image" for artifact in result.artifacts)

    def test_pdf_to_md_omit_image_mode_removes_image_artifacts(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """omit image_mode should leave a placeholder comment and no image artifacts."""
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "page-1.png"))
            return "![[sample_images/page-1.png]]"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": True, "image_mode": "omit"},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "<!-- image omitted: page-1 -->" in md_text
            assert not any(artifact.kind == "image" for artifact in result.artifacts)

    def test_admitted_ofd_format_drives_preprocess_despite_pdf_suffix(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """The admitted format, not a misleading suffix, selects preprocessing."""
        import docwen_plugin_layout.plugin as plugin_module
        from docwen_plugin_layout import LayoutPlugin

        ofd_path = tmp_path / "category_input.pdf"
        ofd_path.write_text("fake ofd content")

        preprocess_calls: list[dict] = []

        def fake_preprocess(input_path, staging_dir, source_format):
            preprocess_calls.append(
                {
                    "input_path": input_path,
                    "source_format": source_format,
                }
            )
            return plugin_module.PreprocessResult(
                effective_input_path=str(sample_pdf_path := tmp_path / "from_ofd.pdf"),
                original_source_format="ofd",
                effective_source_format="pdf",
                intermediate_artifacts=[str(sample_pdf_path)],
            )

        def fake_md_convert(self, context, *, input_path_override=None, source_format_override=None):
            from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
            from docwen_core.models.result import ConversionDiagnostic, ConversionMetrics, ConversionResult

            md_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".md")
            Path(md_path).write_text("# Category test", encoding="utf-8")
            artifact = ArtifactManifest(
                artifact_id="layout-category-md",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=md_path,
                suggested_name="category_input.md",
                media_type="text/markdown",
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

        monkeypatch.setattr(plugin_module, "preprocess_layout_input", fake_preprocess)
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
            assert len(preprocess_calls) == 1
            assert preprocess_calls[0]["source_format"] == "ofd"


class TestPreprocessHelpers:
    def test_layout_category_is_not_resolved_from_file_suffix(self, tmp_path: Path) -> None:
        """A generic category must not bypass admission by guessing a suffix."""
        from docwen_plugin_layout.preprocess import preprocess_layout_input

        pdf_path = tmp_path / "resolved.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")

        result = preprocess_layout_input(str(pdf_path), str(tmp_path), "layout")

        assert result.error_type == "invalid_input"
        assert result.effective_input_path == str(pdf_path)
        assert result.effective_source_format == "layout"
        assert result.original_source_format == "layout"
        assert result.diagnostic_code == "PREPROCESS-SOURCE-FORMAT-NOT-CONCRETE"

    def test_xps_preprocess_opens_explicit_xps_filetype_for_misleading_suffix(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """An admitted XPS is not handed back to PyMuPDF for suffix inference."""
        import fitz

        from docwen_plugin_layout.preprocess import _xps_to_pdf

        misleading_path = tmp_path / "actual-xps-content.pdf"
        create_minimal_xps(misleading_path)
        real_open = fitz.open
        observed_calls: list[tuple[tuple[Any, ...], dict[str, Any]]] = []

        def open_spy(*args: Any, **kwargs: Any):
            observed_calls.append((args, kwargs))
            return real_open(*args, **kwargs)

        monkeypatch.setattr(fitz, "open", open_spy)
        result = _xps_to_pdf(str(misleading_path), str(tmp_path))

        assert result.error_type is None, result.error_message
        assert result.effective_source_format == "pdf"
        assert observed_calls[0] == ((str(misleading_path),), {"filetype": "xps"})
        assert Path(result.effective_input_path).read_bytes().startswith(b"%PDF-")


class TestPreprocessErrors:
    """Preprocess errors must surface as structured ConversionResult errors."""

    def test_ofd_to_md_missing_easyofd(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """OFD→MD with easyofd missing should return dependency_missing."""
        import docwen_plugin_layout.preprocess as pp
        from docwen_plugin_layout import LayoutPlugin

        ofd_path = tmp_path / "test.ofd"
        ofd_path.write_text("fake ofd content")

        # Make easyofd import fail
        def fake_ofd_to_pdf(input_path, staging_dir):
            from docwen_plugin_layout.preprocess import PreprocessResult

            return PreprocessResult(
                effective_input_path=input_path,
                original_source_format="ofd",
                effective_source_format="ofd",
                error_type="dependency_missing",
                error_message="easyofd is not installed",
                diagnostic_code="OFD2PDF-DEPENDENCY-MISSING",
            )

        monkeypatch.setattr(pp, "_ofd_to_pdf", fake_ofd_to_pdf)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(ofd_path),
                staging,
                "md",
                source_format="ofd",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "dependency_missing"
            assert result.error.diagnostic_code == "OFD2PDF-DEPENDENCY-MISSING"

    def test_xps_to_png_missing_pymupdf(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """XPS→PNG with PyMuPDF missing should return dependency_missing."""
        import docwen_plugin_layout.preprocess as pp
        from docwen_plugin_layout import LayoutPlugin

        xps_path = tmp_path / "test.xps"
        xps_path.write_text("fake xps content")

        def fake_xps_to_pdf(input_path, staging_dir):
            from docwen_plugin_layout.preprocess import PreprocessResult

            return PreprocessResult(
                effective_input_path=input_path,
                original_source_format="xps",
                effective_source_format="xps",
                error_type="dependency_missing",
                error_message="PyMuPDF is not installed",
                diagnostic_code="XPS2PDF-DEPENDENCY-MISSING",
            )

        monkeypatch.setattr(pp, "_xps_to_pdf", fake_xps_to_pdf)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(xps_path),
                staging,
                "png",
                source_format="xps",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "dependency_missing"
            assert result.error.diagnostic_code == "XPS2PDF-DEPENDENCY-MISSING"

    def test_ofd_to_pdf_missing_easyofd(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """OFD→PDF (direct route) with easyofd missing should return error."""
        import builtins

        from docwen_plugin_layout import LayoutPlugin

        ofd_path = tmp_path / "test.ofd"
        ofd_path.write_text("fake ofd content")

        # The direct OFD→PDF route uses OfdToPdfConverter which calls
        # LayoutToPdfConverter._ofd_to_pdf() which imports easyofd.
        # Force that import to fail so the test is independent of the
        # environment actually having easyofd installed.
        original_import = builtins.__import__

        def fake_import(name, globals=None, locals=None, fromlist=(), level=0):
            if name == "easyofd":
                raise ImportError("easyofd is not installed")
            return original_import(name, globals, locals, fromlist, level)

        monkeypatch.setattr(builtins, "__import__", fake_import)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(ofd_path),
                staging,
                "pdf",
                source_format="ofd",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"
            assert result.error.diagnostic_code == "OFD2PDF-DEPENDENCY-MISSING"

    def test_ofd_to_pdf_applies_easyofd_patches(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """The direct OFD-to-PDF route applies the shared easyofd patches."""
        import easyofd

        import docwen_core.ofd as ofd_patches
        from docwen_plugin_layout import LayoutPlugin

        ofd_path = tmp_path / "test.ofd"
        ofd_path.write_bytes(b"fake ofd content")
        patch_calls: list[bool] = []

        class _FakeOFD:
            def read(self, path, fmt=None):
                assert patch_calls == [True]

            def to_pdf(self):
                return b"%PDF-1.4 fake ofd output"

        monkeypatch.setattr(ofd_patches, "apply_easyofd_patches", lambda: patch_calls.append(True))
        monkeypatch.setattr(easyofd, "OFD", _FakeOFD)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(ofd_path),
                staging,
                "pdf",
                source_format="ofd",
            )
            result = LayoutPlugin().convert(context)

        assert result.success is True
        assert patch_calls == [True]

    def test_xps_to_pdf_missing_pymupdf(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """XPS→PDF (direct route) with PyMuPDF missing should return error."""
        import docwen_plugin_layout.to_pdf.converter as to_pdf_module
        from docwen_plugin_layout import LayoutPlugin

        xps_path = tmp_path / "test.xps"
        xps_path.write_text("fake xps content")

        # Simulate PyMuPDF not being available for the direct XPS→PDF route

        def fake_convert(self, context):
            from docwen_core.models.result import (
                ConversionDiagnostic,
                ConversionErrorInfo,
                ConversionResult,
            )

            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="dependency_missing",
                    message="PyMuPDF is not installed",
                    diagnostic_code="XPS2PDF-DEPENDENCY-MISSING",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message="PyMuPDF is not installed",
                        code="XPS2PDF-DEPENDENCY-MISSING",
                    )
                ],
            )

        monkeypatch.setattr(to_pdf_module.XpsToPdfConverter, "convert", fake_convert)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(xps_path),
                staging,
                "pdf",
                source_format="xps",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "dependency_missing"
