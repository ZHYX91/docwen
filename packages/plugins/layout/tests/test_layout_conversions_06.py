"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

from ._layout_conversions_support import (
    Any,
    Path,
    _assert_open_pdf_document,
    _build_fake_context,
    _create_text_pdf,
    _ocr_success,
    _page_fragments,
    _write_valid_png,
    pytest,
    sys,
    tempfile,
    types,
)

pytestmark = pytest.mark.contract


class TestPreprocessChain:
    """Verify that OFD/XPS inputs go through the preprocess layer before
    reaching the downstream converters."""

    def test_direct_page_ocr_continues_after_third_page_render_failure(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """A page render failure must not suppress the following physical page."""
        import fitz

        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_core.text.ocr import OcrStatus

        pdf_path = tmp_path / "render-failure.pdf"
        _create_text_pdf(pdf_path, ["ONE", "TWO", "THREE", "FOUR"])
        original_get_pixmap = fitz.Page.get_pixmap
        ocr_calls: list[str] = []

        def selective_get_pixmap(page: Any, *args: Any, **kwargs: Any) -> Any:
            if page.number == 2:
                raise RuntimeError("third-page render failure")
            return original_get_pixmap(page, *args, **kwargs)

        def fake_ocr(path: str, **_kwargs: Any) -> Any:
            ocr_calls.append(Path(path).name)
            return _ocr_success(f"OCR {len(ocr_calls)}")

        monkeypatch.setattr(fitz.Page, "get_pixmap", selective_get_pixmap)
        monkeypatch.setattr(converter_module, "run_ocr_outcome", fake_ocr)

        outcomes = converter_module._ocr_page_outcomes(str(pdf_path), str(tmp_path))

        assert [outcome.status for outcome in outcomes] == [
            OcrStatus.SUCCESS,
            OcrStatus.SUCCESS,
            OcrStatus.RECOGNITION_FAILED,
            OcrStatus.SUCCESS,
        ]
        assert ocr_calls == ["_ocr_page_1.png", "_ocr_page_2.png", "_ocr_page_4.png"]
        assert not list(tmp_path.glob("_ocr_page_*.png"))

    def test_pdf_to_md_shared_extracted_image_is_not_assigned_to_one_page(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """A resource referenced by two pages must remain explicitly unresolved."""
        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "shared-image.pdf"
        _create_text_pdf(pdf_path, ["PAGE 1", "PAGE 2"])

        def fake_to_markdown(
            document: Any,
            *,
            pages: list[int],
            write_images: bool,
            image_path: str,
            **_kwargs: Any,
        ) -> str:
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            shared = Path(image_path) / "shared.png"
            if not shared.exists():
                _write_valid_png(shared)
            return f"![[shared-image_images/shared.png]]\nPAGE {pages[0] + 1}"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(converter_module, "_ocr_page_outcomes", lambda *_args, **_kwargs: [])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": True, "to_md_enable_ocr": False},
            )
            result = LayoutToMarkdownConverter().convert(context)

        assert result.success is True
        image = next(artifact for artifact in result.artifacts if artifact.kind == "image")
        assert "source_page" not in image.metadata
        warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.code == "resource_page_unresolved"]
        assert len(warnings) == 1
        assert warnings[0].artifact_id == image.artifact_id

    def test_pdf_to_md_emits_page_fragment_without_placement_option(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Physical-page OCR has one canonical typed-fragment output."""
        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "page-1.png"))
            return "![[sample_images/page-1.png]]\n\nText"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(
            converter_module,
            "run_ocr_outcome",
            lambda _path, **_kwargs: _ocr_success("INLINE OCR"),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": True, "to_md_enable_ocr": True},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "![[page-1.png]]" in md_text
            assert "INLINE OCR" not in md_text
            assert "## OCR Text" not in md_text
            page_fragment = _page_fragments(result)[0]
            assert Path(page_fragment.staging_path).read_text(encoding="utf-8") == "INLINE OCR\n"

    def test_pdf_to_md_page_fragment_preserves_ocr_text_exactly(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Page OCR is preserved as evidence and never replaced by embedded text."""
        import fitz

        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        document = fitz.open()
        page = document.new_page()
        page.insert_text((72, 72), "武公规〔2022〕2 号", fontname="china-s", fontsize=12)
        document.save(pdf_path)
        document.close()

        def fake_to_markdown(
            document: Any,
            *,
            write_images: bool = False,
            image_path: str | None = None,
            **_kwargs,
        ) -> str:
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "sample.pdf-0001-03.png"))
            return "![[sample_images/sample.pdf-0001-03.png]]"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(
            converter_module,
            "run_ocr_outcome",
            lambda _path, **_kwargs: _ocr_success("武公规\n【2022）\n2\n号"),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": True, "to_md_enable_ocr": True},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "武公规" not in md_text
            fragment_text = Path(_page_fragments(result)[0].staging_path).read_text(encoding="utf-8")
            assert fragment_text == "武公规\n【2022）\n2\n号\n"

    def test_pdf_to_md_image_links_and_page_fragments_are_independent(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Image resources and typed OCR fragments are independent outputs."""
        import docwen_plugin_layout.to_markdown.converter as converter_module
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
        monkeypatch.setattr(
            converter_module,
            "run_ocr_outcome",
            lambda _path, **_kwargs: _ocr_success("SIDECAR OCR"),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={
                    "to_md_keep_images": True,
                    "to_md_enable_ocr": True,
                    "image_link_style": "markdown_embed",
                },
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "sample__img_001_ocr" not in md_text
            assert "page-1.png" in md_text
            assert "SIDECAR OCR" not in md_text

            sidecar = _page_fragments(result)[0]
            sidecar_text = Path(sidecar.staging_path).read_text(encoding="utf-8")
            assert sidecar_text == "SIDECAR OCR\n"

    def test_pdf_to_md_ocr_without_keep_images_emits_only_page_fragment(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """OCR page renders remain temporary when extracted resources are disabled."""
        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is False
            assert image_path == ""
            return "Text"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(
            converter_module,
            "run_ocr_outcome",
            lambda _path, **_kwargs: _ocr_success("ONLY OCR"),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": False, "to_md_enable_ocr": True},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "ONLY OCR" not in md_text
            assert "page-1.png" not in md_text
            assert not any(artifact.kind == "image" for artifact in result.artifacts)
            assert Path(_page_fragments(result)[0].staging_path).read_text(encoding="utf-8") == "ONLY OCR\n"

    def test_direct_page_ocr_helper_preserves_typed_model_failure(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import fitz

        import docwen_plugin_layout.to_markdown.converter as converter_module

        pdf_path = tmp_path / "one-page.pdf"
        doc = fitz.open()
        doc.new_page()
        doc.save(pdf_path)
        doc.close()

        from docwen_core.text.ocr import OcrOutcome, OcrStatus

        def missing_ocr_model(*_args: object, **_kwargs: object) -> OcrOutcome:
            return OcrOutcome(OcrStatus.MODEL_MISSING, message="rapidocr model missing")

        monkeypatch.setattr(converter_module, "run_ocr_outcome", missing_ocr_model)

        outcomes = converter_module._ocr_page_outcomes(str(pdf_path), str(tmp_path))

        assert len(outcomes) == 1
        assert outcomes[0].status is OcrStatus.MODEL_MISSING
        assert outcomes[0].recognized_text == ""

    def test_page_level_ocr_render_failure_preserves_base_conversion(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Base Markdown"])

        def fake_to_markdown(*_args: object, **_kwargs: object) -> str:
            return "Base Markdown"

        def failed_page_ocr(*_args: object, **_kwargs: object) -> list[Any]:
            raise RuntimeError("page render failed")

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(converter_module, "_ocr_page_outcomes", failed_page_ocr)

        fragment_bytes: bytes
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": False, "to_md_enable_ocr": True},
            )
            result = LayoutToMarkdownConverter().convert(context)
            fragment_bytes = Path(_page_fragments(result)[0].staging_path).read_bytes()

        assert result.success is True
        page_fragment = _page_fragments(result)[0]
        assert page_fragment.metadata["ocr_status"] == "recognition_failed"
        assert fragment_bytes == b""
        assert any(
            diagnostic.code == "OCR-BEST-EFFORT" and diagnostic.artifact_id == page_fragment.artifact_id
            for diagnostic in result.diagnostics
        )
        assert any(d.code == "PDF2MD-OCR-OK" for d in result.diagnostics)

    def test_page_level_no_text_warns_about_possible_missed_text(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_core.text.ocr import OcrOutcome, OcrStatus
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Base Markdown"])
        monkeypatch.setitem(
            sys.modules,
            "pymupdf4llm",
            types.SimpleNamespace(to_markdown=lambda *_args, **_kwargs: "Base Markdown"),
        )
        monkeypatch.setattr(
            converter_module,
            "_ocr_page_outcomes",
            lambda *_args, **_kwargs: [OcrOutcome(OcrStatus.NO_TEXT)],
        )

        fragment_bytes: bytes
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": False, "to_md_enable_ocr": True},
            )
            result = LayoutToMarkdownConverter().convert(context)
            fragment_bytes = Path(_page_fragments(result)[0].staging_path).read_bytes()

        assert result.success is True
        page_fragment = _page_fragments(result)[0]
        assert page_fragment.metadata["ocr_status"] == "no_text"
        assert fragment_bytes == b""
        warning = next(diagnostic for diagnostic in result.diagnostics if diagnostic.code == "OCR-BEST-EFFORT")
        assert warning.artifact_id == page_fragment.artifact_id
        assert "status=no_text" in warning.message
        assert "may have been missed" in warning.message
        assert any(d.code == "PDF2MD-OCR-OK" for d in result.diagnostics)

    @pytest.mark.parametrize(
        "status_value",
        ["input_missing", "unavailable", "model_missing", "initialization_failed", "recognition_failed"],
    )
    def test_pdf_to_md_image_operational_ocr_failure_preserves_base_conversion(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        status_value: str,
    ) -> None:
        """Image-level optional OCR errors must not destroy extracted Markdown."""
        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_core.text.ocr import OcrOutcome, OcrStatus
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Base Markdown"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "page-1.png"))
            return "![[sample_images/page-1.png]]"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(
            converter_module,
            "run_ocr_outcome",
            lambda *_args, **_kwargs: OcrOutcome(
                OcrStatus(status_value),
                message="private model path",
            ),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": True, "to_md_enable_ocr": True},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True
            assert result.error is None
            page_fragment = _page_fragments(result)[0]
            warning = next(diagnostic for diagnostic in result.diagnostics if diagnostic.code == "OCR-BEST-EFFORT")
            assert warning.artifact_id == page_fragment.artifact_id
            assert f"status={status_value}" in warning.message
            assert "private" not in warning.message
            markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "page-1.png" in markdown

    def test_pdf_to_md_respects_image_link_style(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """Extracted PDF image refs should use requested Markdown link style."""
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "page-1.png"))
            return "![[sample_images/page-1.png]]\n\nText"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": True, "image_link_style": "markdown_embed"},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "![page-1](page-1.png)" in md_text
            assert "![[sample_images/page-1.png]]" not in md_text
            assert any(artifact.kind == "image" for artifact in result.artifacts)
