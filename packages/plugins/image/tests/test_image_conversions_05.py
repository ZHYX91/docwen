"""Focused tests split from test_image_conversions.py."""

from __future__ import annotations

from ._image_conversions_support import (
    PROJECT_ROOT,
    Any,
    Image,
    Path,
    _build_fake_context,
    json,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.golden


class TestImageMergeToTiff:
    @pytest.mark.contract
    def test_merge_two_images_to_multipage_tiff(self, sample_png_path: Path, sample_second_png_path: Path) -> None:
        from docwen_plugin_image.merge.converter import ImageToTiffMerger

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "tif",
                {"mode": "smart"},
                action_name="merge_images_to_tiff",
                extra_input_paths=[str(sample_second_png_path)],
            )
            result = ImageToTiffMerger().convert(context)

            assert result.success is True
            artifact = result.artifacts[0]
            assert artifact.media_type == "image/tiff"
            assert artifact.metadata["image_count"] == 2
            with Image.open(artifact.staging_path) as img:
                assert img.format == "TIFF"
                assert getattr(img, "n_frames", 1) == 2
            # L-3: metrics
            assert result.metrics.input_bytes > 0
            assert result.metrics.output_bytes > 0
            assert result.metrics.extra["image_count"] == 2

    @pytest.mark.integration
    def test_merge_images_to_tiff_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path) -> None:
        """Merge Images→TIFF should finalize the multipage TIFF into the user output dir."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_image.plugin import ImagePlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_merge_images_to_tiff_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        alpha_scenario = next(
            scenario
            for scenario in fixture["additional_scenarios"]
            if scenario["scenario_id"] == "all_rgba_smart_preserves_alpha"
        )
        scenarios = [
            {
                "scenario_id": "rgb_smart",
                "input_images": fixture["input_images"],
                "expected": fixture["projects"]["docwen-current"],
                "colors": [(220, 20, 20), (20, 60, 220)],
            },
            {
                "scenario_id": "all_rgba_smart_preserves_alpha",
                "input_images": alpha_scenario["input_images"],
                "expected": alpha_scenario["projects"]["docwen-current"],
                "colors": [(220, 20, 20, 96), (20, 60, 220, 160)],
            },
        ]
        registry = PluginRegistry()
        registry.register(ImagePlugin())
        workspace_root = tmp_path / "workspace"
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )

        for scenario in scenarios:
            scenario_id = scenario["scenario_id"]
            scenario_dir = tmp_path / scenario_id
            scenario_dir.mkdir()
            input_paths: list[Path] = []
            for input_info, color in zip(scenario["input_images"], scenario["colors"], strict=True):
                input_path = scenario_dir / input_info["name"]
                Image.new(input_info["mode"], tuple(input_info["size"]), color).save(input_path, format="PNG")
                input_paths.append(input_path)

            output_dir = tmp_path / f"out_{scenario_id}"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id=f"merge-images-to-tiff-finalizer-{scenario_id}",
                input_refs=[
                    FileRef(
                        path=str(path),
                        format="png",
                        category="image",
                        size_bytes=path.stat().st_size,
                    )
                    for path in input_paths
                ],
                target_format="tif",
                action_name="merge_images_to_tiff",
                options={"mode": "smart"},
                output_policy=OutputPolicy(output_dir=str(output_dir)),
            )

            result = task_mgr.execute_single(request)

            assert result.success, f"unexpected error for {scenario_id}: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            expected = scenario["expected"]
            tiff_path = Path(artifact.staging_path)
            assert tiff_path.parent == output_dir
            assert tiff_path.name == expected["suggested_name"]
            assert artifact.media_type == expected["artifact_media_type"]
            assert artifact.metadata["image_count"] == expected["metadata_image_count"]
            assert artifact.metadata["mode"] == expected["metrics_mode"]
            assert tiff_path.suffix.lower() == expected["output_suffix"]
            assert tiff_path.stat().st_size > 0
            assert any(d.code == "IMG2TIFF-OK" for d in result.diagnostics)
            assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
            with Image.open(tiff_path) as img:
                assert img.format == expected["format"]
                assert getattr(img, "n_frames", 1) == expected["n_frames"]
                for index, expected_frame in enumerate(expected["frames"]):
                    img.seek(index)
                    assert img.mode == expected_frame["mode"]
                    assert list(img.size) == expected_frame["size"]
            assert str(workspace_root) not in str(tiff_path)

    @pytest.mark.contract
    def test_merge_all_alpha_smart_mode_preserves_rgba_frames(self, tmp_path: Path) -> None:
        from docwen_plugin_image.merge.converter import ImageToTiffMerger

        first = tmp_path / "alpha-red.png"
        second = tmp_path / "alpha-blue.png"
        Image.new("RGBA", (12, 9), (220, 20, 20, 96)).save(first, format="PNG")
        Image.new("RGBA", (7, 5), (20, 60, 220, 160)).save(second, format="PNG")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(first),
                staging,
                "tif",
                {"mode": "smart"},
                action_name="merge_images_to_tiff",
                extra_input_paths=[str(second)],
            )
            result = ImageToTiffMerger().convert(context)

            assert result.success is True
            with Image.open(result.artifacts[0].staging_path) as img:
                assert img.format == "TIFF"
                assert getattr(img, "n_frames", 1) == 2
                img.seek(0)
                assert img.mode == "RGBA"
                assert img.size == (12, 9)
                img.seek(1)
                assert img.mode == "RGBA"
                assert img.size == (7, 5)
            assert result.metrics.extra["mode"] == "smart"

    @pytest.mark.contract
    def test_merge_single_image_creates_single_frame_tiff(self, sample_png_path: Path) -> None:
        """M-3: merging a single image should produce a single-frame TIFF."""
        from docwen_plugin_image.merge.converter import ImageToTiffMerger

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "tif",
                {"mode": "smart"},
                action_name="merge_images_to_tiff",
            )
            result = ImageToTiffMerger().convert(context)

            assert result.success is True
            artifact = result.artifacts[0]
            assert artifact.media_type == "image/tiff"
            assert artifact.metadata["image_count"] == 1
            with Image.open(artifact.staging_path) as img:
                assert img.format == "TIFF"
                assert getattr(img, "n_frames", 1) == 1

    @pytest.mark.contract
    def test_merge_with_non_image_input_fails(self, sample_png_path: Path, tmp_path: Path) -> None:
        """M-3: mixing a non-image file should cause conversion_failed error."""
        from docwen_plugin_image.merge.converter import ImageToTiffMerger

        # Create a fake .txt file disguised as an image in the input_refs
        bad_path = tmp_path / "notanimage.png"
        bad_path.write_text("This is just text, not an image.", encoding="utf-8")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "tif",
                {"mode": "smart"},
                action_name="merge_images_to_tiff",
                extra_input_paths=[str(bad_path)],
            )
            result = ImageToTiffMerger().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"
            assert result.error.diagnostic_code == "IMG2TIFF-ERROR"
            assert "notanimage.png" in result.error.message
            assert "Failed to load image" in result.error.message


@pytest.mark.contract
class TestAdmittedFormatRouting:
    """Parser selection consumes the admitted concrete format, not the suffix."""

    def test_to_markdown_projects_detected_png_for_txt_suffix(self, tmp_path: Path) -> None:
        from docwen_plugin_image import ImagePlugin

        input_path = tmp_path / "misnamed.txt"
        Image.new("RGB", (12, 8), (20, 120, 200)).save(input_path, format="PNG")
        staging = tmp_path / "staging-md"
        staging.mkdir()
        context = _build_fake_context(
            str(input_path),
            str(staging),
            "md",
            {"image_mode": "file", "to_md_keep_images": True, "to_md_enable_ocr": False},
            source_format="png",
        )

        result = ImagePlugin().convert(context)

        assert result.success is True
        image_artifact = next(artifact for artifact in result.artifacts if artifact.kind == "image")
        markdown_artifact = next(artifact for artifact in result.artifacts if artifact.kind == "primary")
        assert image_artifact.suggested_name == "misnamed.png"
        assert Path(image_artifact.staging_path).suffix == ".png"
        assert image_artifact.media_type == "image/png"
        markdown = Path(markdown_artifact.staging_path).read_text(encoding="utf-8")
        assert "misnamed.png" in markdown

    def test_format_conversion_splits_detected_multipage_tiff_with_png_suffix(self, tmp_path: Path) -> None:
        from docwen_plugin_image import ImagePlugin

        input_path = tmp_path / "multipage.png"
        first = Image.new("RGB", (10, 7), (255, 0, 0))
        second = Image.new("RGB", (10, 7), (0, 255, 0))
        first.save(input_path, format="TIFF", save_all=True, append_images=[second])
        first.close()
        second.close()
        staging = tmp_path / "staging-format"
        staging.mkdir()
        context = _build_fake_context(
            str(input_path),
            str(staging),
            "png",
            source_format="tiff",
        )

        result = ImagePlugin().convert(context)

        assert result.success is True
        assert len(result.artifacts) == 2
        assert result.metrics.extra["frame_count"] == 2
        assert [artifact.metadata["source_format"] for artifact in result.artifacts] == ["tif", "tif"]
        assert [artifact.suggested_name for artifact in result.artifacts] == [
            "multipage_page1.png",
            "multipage_page2.png",
        ]

    def test_to_pdf_uses_detected_bmp_despite_png_suffix(self, tmp_path: Path) -> None:
        from docwen_plugin_image import ImagePlugin

        input_path = tmp_path / "bitmap.png"
        Image.new("RGB", (14, 9), (10, 40, 90)).save(input_path, format="BMP")
        staging = tmp_path / "staging-pdf"
        staging.mkdir()
        context = _build_fake_context(
            str(input_path),
            str(staging),
            "pdf",
            source_format="bmp",
        )

        result = ImagePlugin().convert(context)

        assert result.success is True
        assert len(result.artifacts) == 1
        assert Path(result.artifacts[0].staging_path).read_bytes().startswith(b"%PDF")

    def test_png_with_heic_suffix_does_not_enter_heic_preconversion(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_image._common as common
        from docwen_plugin_image import ImagePlugin

        input_path = tmp_path / "ordinary.heic"
        Image.new("RGB", (9, 6), (100, 60, 20)).save(input_path, format="PNG")
        staging = tmp_path / "staging-no-heic"
        staging.mkdir()

        def unexpected_preconversion(_input_path: str, _staging_dir: str) -> str:
            raise AssertionError("suffix incorrectly selected HEIC preprocessing")

        monkeypatch.setattr(common, "preconvert_heic_to_png", unexpected_preconversion)
        context = _build_fake_context(
            str(input_path),
            str(staging),
            "jpg",
            source_format="png",
        )

        result = ImagePlugin().convert(context)

        assert result.success is True
        assert result.artifacts[0].media_type == "image/jpeg"


@pytest.mark.contract
class TestErrorPaths:
    def test_cancellation_before_format_conversion_is_raised(self, sample_png_path: Path) -> None:
        """L-2: Pre-cancelled token should cause CancellationRequested."""
        from docwen_core.errors import CancellationRequested
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_png_path), staging, "jpg", pre_cancelled=True)
            with pytest.raises(CancellationRequested):
                ImageFormatConverter().convert(context)

    def test_corrupt_png_returns_conversion_failed(self, corrupt_png_path: Path) -> None:
        """M-2: corrupt/truncated image should return success=False with conversion_failed."""
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(corrupt_png_path), staging, "jpg")
            result = ImageFormatConverter().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"
            assert result.error.diagnostic_code == "IMAGEFMT-CONVERT-ERROR"

    def test_non_image_file_disguised_as_png_returns_error(self, fake_txt_as_png_path: Path) -> None:
        """M-2: text file disguised as .png should fail with conversion error."""
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(fake_txt_as_png_path), staging, "jpg")
            result = ImageFormatConverter().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"

    def test_corrupt_image_to_markdown_file_mode_handles_gracefully(self, corrupt_png_path: Path) -> None:
        """M-2: corrupt image in to_markdown file-mode should not crash.

        The to_markdown converter does not open the image with PIL — it copies
        the file as-is. A corrupt-but-valid-extension file is copied successfully
        (the consumer of the Markdown output is responsible for image validity).

        OCR is disabled here because a corrupt image may cause RapidOCR to fail
        in unpredictable ways (the test verifies the copy path, not OCR).
        """
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(corrupt_png_path), staging, "md", {"to_md_enable_ocr": False})
            result = ImageToMarkdownConverter().convert(context)

            # file-mode copies without PIL validation — succeeds gracefully
            assert result.success is True
            assert len(result.artifacts) == 2  # MD artifact + image artifact


@pytest.mark.contract
class TestHeicPreprocessing:
    def test_heic_preconversion_dispatches_with_narrow_hub_context(
        self,
        real_heic_path: Path,
        monkeypatch,
    ) -> None:
        """The PNG intermediate should require only the internal converter contract."""
        from docwen_core.models.result import ConversionResult
        from docwen_core.protocols.hub_context import HubConversionContext
        from docwen_plugin_image import ImagePlugin
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        received_contexts: list[Any] = []

        def _capture_context(_converter: ImageFormatConverter, context: Any) -> ConversionResult:
            received_contexts.append(context)
            return ConversionResult(task_id=context.request.request_id, success=True)

        monkeypatch.setattr(ImageFormatConverter, "convert", _capture_context)
        plugin = ImagePlugin()
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(real_heic_path), staging, "jpg")

            result = plugin.convert(context)

        assert result.success is True
        assert len(received_contexts) == 1
        downstream_context = received_contexts[0]
        assert isinstance(downstream_context, HubConversionContext)
        assert Path(downstream_context.workspace.input_path).suffix == ".png"
        assert not hasattr(downstream_context, "numbering_registry")
        assert not hasattr(downstream_context, "proofread_rules")

    def test_heic_input_preconverts_to_png_then_uses_existing_converter(self, real_heic_path: Path) -> None:
        """HEIC input restores the old optional pillow-heif preprocessing path."""
        from docwen_plugin_image import ImagePlugin

        plugin = ImagePlugin()
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(real_heic_path), staging, "jpg")
            result = plugin.convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.suggested_name == "sample.jpg"
            assert artifact.media_type == "image/jpeg"
            with Image.open(artifact.staging_path) as img:
                assert img.format == "JPEG"
                assert img.size == (18, 12)

    def test_heic_input_preconverts_before_pdf_conversion(self, real_heic_path: Path) -> None:
        """HEIC preprocessing should feed the existing image-to-PDF route."""
        from docwen_plugin_image import ImagePlugin

        plugin = ImagePlugin()
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(real_heic_path), staging, "pdf", {"quality_mode": "original"})
            result = plugin.convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.media_type == "application/pdf"
            assert artifact.suggested_name == "sample.pdf"
            assert Path(artifact.staging_path).read_bytes().startswith(b"%PDF")

    def test_heif_input_preconverts_to_png_for_markdown_assets(self, real_heif_path: Path) -> None:
        """HEIF-to-Markdown should emit renderable PNG assets after preprocessing."""
        from docwen_plugin_image import ImagePlugin

        plugin = ImagePlugin()
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(real_heif_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_keep_images": True, "to_md_enable_ocr": False},
            )
            result = plugin.convert(context)

            assert result.success is True
            assert len(result.artifacts) == 2
            md_artifact = result.artifacts[0]
            image_artifact = result.artifacts[1]
            assert md_artifact.media_type == "text/markdown"
            assert image_artifact.media_type == "image/png"
            assert image_artifact.suggested_name == "sample.png"
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "source_format: png" in md_text
            assert "![[sample.png]]" in md_text

    def test_invalid_heic_input_returns_preprocess_failure(self, sample_heic_path: Path) -> None:
        """Invalid HEIC content should fail during preprocessing with a clear diagnostic."""
        from docwen_plugin_image import ImagePlugin

        plugin = ImagePlugin()
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_heic_path), staging, "png")
            result = plugin.convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"
            assert result.error.diagnostic_code == "IMG-HEIC-PREPROCESS-ERROR"
            assert any(d.code == "IMG-HEIC-PREPROCESS-ERROR" for d in result.diagnostics)
