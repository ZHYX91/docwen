"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

from ._layout_conversions_support import (
    Any,
    Path,
    _assert_open_pdf_document,
    _build_fake_context,
    _build_runtime_pipeline,
    _create_text_pdf,
    _document_node_root,
    _io_path,
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

    def test_pdf_to_md_real_page_level_ocr_smoke(self, tmp_path: Path) -> None:
        """Real OCR smoke: every physical page becomes its own OCR fragment."""
        import fitz

        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "layout_ocr_smoke.pdf"
        doc = fitz.open()
        page = doc.new_page(width=720, height=220)
        page.insert_text((36, 112), "HELLO DOCWEN OCR", fontsize=54)
        doc.save(str(pdf_path))
        doc.close()

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                source_format="pdf",
                options={"to_md_keep_images": False, "to_md_enable_ocr": True, "render_dpi": 220},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "# HELLO DOCWEN OCR" in md_text
            assert "## OCR Text" not in md_text
            page_fragments = _page_fragments(result)
            assert len(page_fragments) == 1
            assert "HELLO DOCWEN OCR" in Path(page_fragments[0].staging_path).read_text(encoding="utf-8")
            assert page_fragments[0].metadata == {
                "fragment_kind": "page",
                "page_index": 1,
                "page_count": 1,
                "source_page": 1,
                "ocr_status": "success",
            }
            assert result.metrics.extra["ocr_enabled"] is True
            assert result.metrics.extra["ocr_pages"] == 1
            assert result.metrics.extra["ocr_images"] == 0
            assert any(diagnostic.code == "PDF2MD-OCR-OK" for diagnostic in result.diagnostics)

    def test_pdf_to_md_real_scanned_page_ocr_smoke(self, tmp_path: Path) -> None:
        """Real OCR smoke: an image-only page still emits one page fragment."""
        import fitz
        from PIL import Image, ImageDraw, ImageFont

        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        scan_image = Image.new("RGB", (1400, 420), "white")
        draw = ImageDraw.Draw(scan_image)
        try:
            font = ImageFont.truetype("DejaVuSans.ttf", 86)
        except OSError:
            font = ImageFont.load_default()
        draw.text((70, 140), "SCANNED DOCWEN OCR", fill="black", font=font)

        scan_path = tmp_path / "scan-page.png"
        scan_image.save(scan_path)
        pdf_path = tmp_path / "scanned-layout-ocr-smoke.pdf"
        doc = fitz.open()
        page = doc.new_page(width=700, height=210)
        page.insert_image(fitz.Rect(0, 0, 700, 210), filename=str(scan_path))
        doc.save(str(pdf_path))
        doc.close()

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                source_format="pdf",
                options={"to_md_keep_images": False, "to_md_enable_ocr": True, "render_dpi": 260},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "# SCANNED DOCWEN OCR" not in md_text
            assert "## OCR Text" not in md_text
            page_fragments = _page_fragments(result)
            assert len(page_fragments) == 1
            assert "SCANNED DOCWEN OCR" in Path(page_fragments[0].staging_path).read_text(encoding="utf-8")
            assert page_fragments[0].metadata["source_page"] == 1
            assert result.metrics.extra["ocr_enabled"] is True
            assert result.metrics.extra["ocr_pages"] == 1
            assert result.metrics.extra["ocr_images"] == 0
            assert any(diagnostic.code == "PDF2MD-OCR-OK" for diagnostic in result.diagnostics)

    def test_pdf_to_md_real_image_file_smoke(self, tmp_path: Path) -> None:
        """Real pymupdf4llm smoke: extracted image refs should become artifacts."""
        import fitz
        from PIL import Image

        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        source_img = tmp_path / "source.png"
        Image.new("RGB", (20, 20), (255, 0, 0)).save(source_img)
        pdf_path = tmp_path / "image-source.pdf"
        doc = fitz.open()
        page = doc.new_page(width=200, height=200)
        page.insert_text((20, 20), "Image PDF", fontsize=12)
        page.insert_image(fitz.Rect(20, 40, 80, 100), filename=str(source_img))
        doc.save(str(pdf_path))
        doc.close()

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                source_format="pdf",
                options={"to_md_keep_images": True, "to_md_enable_ocr": False, "image_link_style": "markdown_embed"},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "Image PDF" in md_text
            assert "](" in md_text
            assert str(tmp_path) not in md_text
            image_artifacts = [artifact for artifact in result.artifacts if artifact.kind == "image"]
            assert len(image_artifacts) == 1
            assert result.metrics.extra["image_count"] == 1

    def test_pdf_to_md_real_image_base64_smoke(self, tmp_path: Path) -> None:
        """Real pymupdf4llm smoke: base64 mode should inline extracted images."""
        import fitz
        from PIL import Image

        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        source_img = tmp_path / "source.png"
        Image.new("RGB", (20, 20), (0, 0, 255)).save(source_img)
        pdf_path = tmp_path / "image-source.pdf"
        doc = fitz.open()
        page = doc.new_page(width=200, height=200)
        page.insert_text((20, 20), "Base64 PDF", fontsize=12)
        page.insert_image(fitz.Rect(20, 40, 80, 100), filename=str(source_img))
        doc.save(str(pdf_path))
        doc.close()

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                source_format="pdf",
                options={
                    "to_md_keep_images": True,
                    "to_md_enable_ocr": False,
                    "image_mode": "base64",
                    "image_link_style": "markdown_embed",
                },
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "Base64 PDF" in md_text
            assert "data:image/png;base64," in md_text
            assert not any(artifact.kind == "image" for artifact in result.artifacts)
            assert result.metrics.extra["image_count"] == 0

    def test_pdf_to_md_real_image_and_page_fragment_smoke(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Image extraction and physical-page OCR remain independent outputs."""
        import fitz
        from PIL import Image

        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        source_img = tmp_path / "source.png"
        Image.new("RGB", (20, 20), (0, 255, 0)).save(source_img)
        pdf_path = tmp_path / "image-source.pdf"
        doc = fitz.open()
        page = doc.new_page(width=200, height=200)
        page.insert_text((20, 20), "Sidecar PDF", fontsize=12)
        page.insert_image(fitz.Rect(20, 40, 80, 100), filename=str(source_img))
        doc.save(str(pdf_path))
        doc.close()

        ocr_calls: list[tuple[str, str, str]] = []

        def fake_ocr(
            _path: str,
            *,
            source_format: str,
            ocr_language: str | None = None,
            current_locale: str = "zh_CN",
        ) -> Any:
            ocr_calls.append((source_format, ocr_language or "", current_locale))
            return _ocr_success("REAL IMAGE OCR")

        monkeypatch.setattr(converter_module, "run_ocr_outcome", fake_ocr)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                source_format="pdf",
                options={
                    "to_md_keep_images": True,
                    "to_md_enable_ocr": True,
                    "image_link_style": "markdown_embed",
                    "ocr_language": "english",
                    "locale": "en_US",
                },
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert ocr_calls == [("png", "english", "en_US")]
            md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "Sidecar PDF" in md_text
            assert "image-source__img_001_ocr" not in md_text
            assert "REAL IMAGE OCR" not in md_text

            image_artifacts = [artifact for artifact in result.artifacts if artifact.kind == "image"]
            sidecar_artifacts = _page_fragments(result)
            assert len(image_artifacts) == 1
            assert len(sidecar_artifacts) == 1
            sidecar_text = Path(sidecar_artifacts[0].staging_path).read_text(encoding="utf-8")
            assert sidecar_text == "REAL IMAGE OCR\n"
            assert sidecar_artifacts[0].metadata["source_page"] == 1
            assert image_artifacts[0].metadata["source_page"] == 1
            assert result.metrics.extra["ocr_images"] == 0

    def test_pdf_to_md_image_sidecar_artifacts_are_finalized_through_runtime(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Image and physical-page artifacts must survive runtime finalization."""
        import shutil

        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        pdf_path = tmp_path / "layout-image-artifact-probe.pdf"
        _create_text_pdf(pdf_path, ["Runtime image text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            assert write_images is True
            assert image_path is not None
            Path(image_path).mkdir(parents=True, exist_ok=True)
            _write_valid_png(Path(image_path, "page-1.png"))
            return "![[layout-image-artifact-probe_images/page-1.png]]\n\nRuntime image text"

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(
            converter_module,
            "run_ocr_outcome",
            lambda _path, **_kwargs: _ocr_success("RUNTIME IMAGE OCR"),
        )

        task_manager, ws_manager, ws_root = _build_runtime_pipeline()
        try:
            output_dir = tmp_path / "output_layout_images"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id="layout-image-finalize",
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
                    "to_md_keep_images": True,
                    "to_md_enable_ocr": True,
                    "image_link_style": "markdown_embed",
                },
            )

            result = task_manager.execute_single(request)
        finally:
            ws_manager.cleanup_all()
            shutil.rmtree(ws_root, ignore_errors=True)

        assert result.success is True, f"unexpected error: {result.error}"
        markdown_artifact = next(artifact for artifact in result.artifacts if artifact.is_primary)
        image_artifacts = [artifact for artifact in result.artifacts if artifact.kind == "image"]
        sidecar_artifacts = _page_fragments(result)
        assert len(image_artifacts) == 1
        assert len(sidecar_artifacts) == 1
        node_root = _document_node_root(Path(markdown_artifact.staging_path), output_dir)
        assert _document_node_root(Path(image_artifacts[0].staging_path), output_dir) == node_root
        assert _document_node_root(Path(sidecar_artifacts[0].staging_path), output_dir) == node_root
        assert Path(image_artifacts[0].staging_path).is_file()
        assert _io_path(sidecar_artifacts[0].staging_path).is_file()

        md_text = Path(markdown_artifact.staging_path).read_text(encoding="utf-8")
        sidecar_text = _io_path(sidecar_artifacts[0].staging_path).read_text(encoding="utf-8")
        assert sidecar_artifacts[0].suggested_name not in md_text
        assert image_artifacts[0].suggested_name in md_text
        assert "RUNTIME IMAGE OCR" not in md_text
        assert sidecar_text == "RUNTIME IMAGE OCR\n"
        assert str(output_dir) not in md_text + sidecar_text
        assert "docwen_layout_runtime_" not in md_text + sidecar_text

    def test_pdf_to_md_ocr_uses_physical_pages_even_when_images_are_extracted(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Extracted image count must never replace physical-page OCR."""
        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "sample.pdf"
        _create_text_pdf(pdf_path, ["Text"])

        def fake_to_markdown(document: Any, *, write_images: bool = False, image_path: str | None = None, **_kwargs):
            _assert_open_pdf_document(document, pdf_path)
            if write_images and image_path:
                Path(image_path).mkdir(parents=True, exist_ok=True)
                _write_valid_png(Path(image_path, "page-1.png"))
                return "![[sample_images/page-1.png]]\n\nText"
            return "Text"

        page_ocr_calls: list[int] = []

        def fake_page_ocr(*_args, **_kwargs):
            page_ocr_calls.append(1)
            return [_ocr_success("PAGE OCR")]

        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(converter_module, "_ocr_page_outcomes", fake_page_ocr)
        monkeypatch.setattr(converter_module, "run_ocr_outcome", lambda *_args, **_kwargs: pytest.fail("wrong path"))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": True, "to_md_enable_ocr": True},
            )
            result = LayoutToMarkdownConverter().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            md_artifact = next(artifact for artifact in result.artifacts if artifact.kind == "primary")
            page_artifact = next(
                artifact for artifact in result.artifacts if artifact.metadata.get("fragment_kind") == "page"
            )
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "PAGE OCR" not in md_text
            assert "## OCR Text" not in md_text
            assert "sample_images/" not in md_text
            assert md_text.count("page-1.png") == 1
            assert Path(page_artifact.staging_path).read_text(encoding="utf-8") == "PAGE OCR\n"
            assert page_artifact.metadata == {
                "fragment_kind": "page",
                "page_index": 1,
                "page_count": 1,
                "source_page": 1,
                "ocr_status": "success",
            }
            image_artifact = next(artifact for artifact in result.artifacts if artifact.kind == "image")
            assert image_artifact.metadata["source_page"] == 1
            assert page_ocr_calls == [1]

    @pytest.mark.parametrize(
        ("enable_ocr", "keep_images", "expected_fragments", "expected_images"),
        [(False, False, 0, 0), (True, False, 4, 0), (False, True, 0, 5), (True, True, 4, 5)],
    )
    def test_pdf_to_md_physical_page_matrix_keeps_p_and_k_independent(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        enable_ocr: bool,
        keep_images: bool,
        expected_fragments: int,
        expected_images: int,
    ) -> None:
        """Canonical P=4/K=5 producer corpus freezes all four option combinations."""
        import docwen_plugin_layout.to_markdown.converter as converter_module
        from docwen_application.bundle_mapping import build_bundle_draft
        from docwen_core.models import validate_artifact_bundle_draft
        from docwen_core.text.ocr import OcrOutcome, OcrStatus
        from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter

        pdf_path = tmp_path / "physical-matrix.pdf"
        _create_text_pdf(pdf_path, ["NATIVE 1", "NATIVE 2", "NATIVE 3", "NATIVE 4"])

        def fake_to_markdown(
            document: Any,
            *,
            pages: list[int],
            write_images: bool = False,
            image_path: str = "",
            **_kwargs: Any,
        ) -> str:
            _assert_open_pdf_document(document, pdf_path)
            assert len(pages) == 1
            page_number = pages[0] + 1
            links: list[str] = []
            if write_images:
                image_dir = Path(image_path)
                image_dir.mkdir(parents=True, exist_ok=True)
                image_name = f"page-{page_number}.png"
                _write_valid_png(image_dir / image_name)
                links.append(f"![[physical-matrix_images/{image_name}]]")
                if page_number == 1:
                    links.append("![[physical-matrix_images/unresolved.png]]")
            return "\n".join([*links, f"NATIVE {page_number}"])

        statuses = (
            OcrOutcome(OcrStatus.SUCCESS, text="OCR PAGE 1"),
            OcrOutcome(OcrStatus.NO_TEXT),
            OcrOutcome(OcrStatus.RECOGNITION_FAILED, message="private failure"),
            OcrOutcome(OcrStatus.SUCCESS, text="OCR PAGE 4"),
        )
        monkeypatch.setitem(sys.modules, "pymupdf4llm", types.SimpleNamespace(to_markdown=fake_to_markdown))
        monkeypatch.setattr(converter_module, "_ocr_page_outcomes", lambda *_args, **_kwargs: list(statuses))

        with tempfile.TemporaryDirectory() as staging:
            if keep_images:
                _write_valid_png(Path(staging) / "physical-matrix_images" / "unresolved.png")
            context = _build_fake_context(
                str(pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": keep_images, "to_md_enable_ocr": enable_ocr},
            )
            result = LayoutToMarkdownConverter().convert(context)
            primary_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            fragments = _page_fragments(result)
            images = [artifact for artifact in result.artifacts if artifact.kind == "image"]
            draft = build_bundle_draft(
                profile="physical_page_ocr",
                output_media_type="text/markdown",
                artifacts=result.artifacts,
            )
            validate_artifact_bundle_draft(draft)

        assert result.success is True
        assert len(result.artifacts) == 1 + expected_fragments + expected_images
        assert len(fragments) == expected_fragments
        assert len(images) == expected_images
        assert "OCR PAGE 1" not in primary_text
        assert "OCR PAGE 4" not in primary_text
        if enable_ocr:
            assert [artifact.metadata["ocr_status"] for artifact in fragments] == [
                "success",
                "no_text",
                "recognition_failed",
                "success",
            ]
        if keep_images:
            resolved = [artifact for artifact in images if "source_page" in artifact.metadata]
            unresolved = [artifact for artifact in images if "source_page" not in artifact.metadata]
            assert [artifact.metadata["source_page"] for artifact in resolved] == [1, 2, 3, 4]
            assert len(unresolved) == 1
            warnings = [
                diagnostic for diagnostic in result.diagnostics if diagnostic.code == "resource_page_unresolved"
            ]
            assert [diagnostic.artifact_id for diagnostic in warnings] == [unresolved[0].artifact_id]
