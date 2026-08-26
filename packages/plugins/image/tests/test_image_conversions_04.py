"""Focused tests split from test_image_conversions.py."""

from __future__ import annotations

from ._image_conversions_support import (
    Any,
    Path,
    _build_fake_context,
    _ocr_success,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.golden


class TestImageToMarkdown:
    @pytest.mark.contract
    def test_image_to_markdown_ocr_success_warns_and_preserves_recognized_text(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """OCR success path should append recognised text to the Markdown output."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        calls: list[tuple[str, str, str]] = []

        def fake_ocr(
            _path: str,
            *,
            source_format: str,
            ocr_language: str | None = None,
            current_locale: str = "zh_CN",
        ) -> Any:
            calls.append((source_format, ocr_language or "", current_locale))
            return _ocr_success("识别文本\n第二行")

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", fake_ocr)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "to_md_enable_ocr": True,
                    "to_md_keep_images": True,
                    "ocr_placement": "main_md",
                    "ocr_language": "japanese",
                    "locale": "ja_JP",
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert calls == [("png", "japanese", "ja_JP")]
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "![[sample.png]]" in md_text
            assert "> 识别文本" in md_text
            assert "> 第二行" in md_text
            assert result.metrics.extra["ocr_enabled"] is True
            assert result.metrics.extra["ocr_chars"] == len("识别文本\n第二行")
            assert any(d.code == "IMG2MD-OCR-OK" for d in result.diagnostics)
            assert len(context.progress.diagnostics) == 1
            warning = context.progress.diagnostics[0]
            assert warning[2] == "OCR-BEST-EFFORT"
            assert "status=success" in warning[1]
            assert "may contain recognition errors or omissions" in warning[1]

    @pytest.mark.contract
    def test_image_to_markdown_ocr_disabled_skips_ocr(self, sample_png_path: Path) -> None:
        """When OCR is explicitly disabled, no OCR error should occur."""
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path), staging, "md", {"to_md_enable_ocr": False, "to_md_keep_images": True}
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 2
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "source_format: png" in md_text
            assert "![[sample.png]]" in md_text
            assert result.metrics.extra.get("ocr_enabled") is False

    @pytest.mark.contract
    def test_image_to_markdown_consumes_configured_ocr_default(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Runtime config should control image OCR when request options omit it."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        calls: list[str] = []

        def fake_ocr(_path: str, **_kwargs: object) -> Any:
            calls.append("called")
            return _ocr_success("SHOULD NOT RUN")

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", fake_ocr)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {},
                config_values={"image": {"to_md_enable_ocr": False, "to_md_keep_images": True}},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert calls == []
            assert result.metrics.extra.get("ocr_enabled") is False

    @pytest.mark.contract
    def test_image_to_markdown_options_override_configured_ocr_default(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Explicit request options should stay higher priority than config."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        calls: list[tuple[str, str | None, str]] = []

        def fake_ocr(
            _path: str,
            *,
            source_format: str,
            ocr_language: str | None = None,
            current_locale: str = "zh_CN",
        ) -> Any:
            calls.append((source_format, ocr_language, current_locale))
            return _ocr_success("OVERRIDE OCR")

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", fake_ocr)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"to_md_enable_ocr": True, "ocr_language": "japanese", "ocr_placement": "main_md"},
                config_values={"image": {"to_md_enable_ocr": False, "ocr_language": "english"}},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert calls == [("png", "japanese", "zh_CN")]
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "> OVERRIDE OCR" in md_text
            assert result.metrics.extra.get("ocr_enabled") is True

    @pytest.mark.contract
    def test_keep_images_false_with_file_mode_produces_omit_comment(self, sample_png_path: Path) -> None:
        """H-2: keep_images=False + image_mode="file" triggers omit comment path.

        The combined condition ``(not keep_images and image_mode != "base64")``
        must produce an HTML comment placeholder with no image artifact registered.
        """
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_keep_images": False, "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            # Only the MD artifact — no image artifact should be registered
            assert len(result.artifacts) == 1
            assert result.artifacts[0].media_type == "text/markdown"
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "<!-- image omitted: sample.png -->" in md_text
            assert "![sample]" not in md_text
            assert result.metrics.extra["artifact_count"] == 1


class TestTiffToMarkdown:
    @pytest.mark.contract
    def test_tiff_to_markdown_emits_one_fragment_per_frame_and_continues_after_ocr_failure(
        self,
        sample_four_frame_tiff_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_core.text.ocr import OcrOutcome, OcrStatus
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        calls: list[Path] = []

        def fake_ocr(path: str, **_kwargs: object) -> OcrOutcome:
            calls.append(Path(path))
            index = len(calls)
            if index == 1:
                return OcrOutcome(OcrStatus.SUCCESS, text="PAGE 1")
            if index == 2:
                return OcrOutcome(OcrStatus.NO_TEXT)
            if index == 3:
                raise RuntimeError("private OCR failure")
            return OcrOutcome(OcrStatus.SUCCESS, text="PAGE 4")

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", fake_ocr)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_four_frame_tiff_path),
                staging,
                "md",
                {"to_md_enable_ocr": True, "to_md_keep_images": False},
                source_format="tif",
            )
            result = ImageToMarkdownConverter().convert(context)
            primary_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            fragments = [artifact for artifact in result.artifacts if artifact.kind == "auxiliary"]
            fragment_bytes = [Path(artifact.staging_path).read_bytes() for artifact in fragments]
            private_pngs = list(Path(staging).glob("*.png"))

        assert result.success is True
        assert len(calls) == 4
        assert len({path.name for path in calls}) == 4
        assert [artifact.metadata["ocr_status"] for artifact in fragments] == [
            "success",
            "no_text",
            "recognition_failed",
            "success",
        ]
        assert fragment_bytes == [b"PAGE 1\n", b"", b"", b"PAGE 4\n"]
        assert "PAGE 1" not in primary_text
        assert "PAGE 4" not in primary_text
        assert private_pngs == []
        warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.code == "OCR-BEST-EFFORT"]
        assert len(warnings) == 4
        assert all(diagnostic.artifact_id is not None for diagnostic in warnings)
        assert all("private" not in diagnostic.message for diagnostic in warnings)

    @pytest.mark.parametrize(
        ("enable_ocr", "keep_images", "expected_fragments", "expected_images"),
        [(False, False, 0, 0), (True, False, 4, 0), (False, True, 0, 4), (True, True, 4, 4)],
    )
    @pytest.mark.contract
    def test_tiff_to_markdown_four_option_combinations_are_independent(
        self,
        sample_four_frame_tiff_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        enable_ocr: bool,
        keep_images: bool,
        expected_fragments: int,
        expected_images: int,
    ) -> None:
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        calls: list[str] = []

        def fake_ocr(path: str, **_kwargs: object) -> Any:
            calls.append(path)
            return _ocr_success(f"PAGE {len(calls)}")

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", fake_ocr)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_four_frame_tiff_path),
                staging,
                "md",
                {"to_md_enable_ocr": enable_ocr, "to_md_keep_images": keep_images},
                source_format="tif",
            )
            result = ImageToMarkdownConverter().convert(context)
            fragments = [artifact for artifact in result.artifacts if artifact.kind == "auxiliary"]
            images = [artifact for artifact in result.artifacts if artifact.kind == "image"]

        assert result.success is True
        assert len(result.artifacts) == 1 + expected_fragments + expected_images
        assert len(fragments) == expected_fragments
        assert len(images) == expected_images
        assert len(calls) == (4 if enable_ocr else 0)
        assert [artifact.metadata["source_page"] for artifact in images] == list(range(1, expected_images + 1))

    @pytest.mark.contract
    def test_tiff_to_markdown_legacy_ocr_placement_does_not_change_fragment_granularity(
        self,
        sample_four_frame_tiff_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", lambda *_args, **_kwargs: _ocr_success("OCR"))
        shapes: list[list[tuple[str, dict[str, Any]]]] = []
        for placement in ("main_md", "image_md"):
            with tempfile.TemporaryDirectory() as staging:
                context = _build_fake_context(
                    str(sample_four_frame_tiff_path),
                    staging,
                    "md",
                    {
                        "to_md_enable_ocr": True,
                        "to_md_keep_images": True,
                        "ocr_placement": placement,
                    },
                    source_format="tif",
                )
                result = ImageToMarkdownConverter().convert(context)
                shapes.append([(artifact.kind, artifact.metadata) for artifact in result.artifacts])

        assert shapes[0] == shapes[1]

    @pytest.mark.contract
    def test_tiff_to_markdown_cancellation_between_frames_removes_private_outputs(
        self,
        sample_four_frame_tiff_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_core.cancellation import CancellationToken
        from docwen_core.errors import CancellationRequested
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        token = CancellationToken()

        def cancel_after_first(_path: str, **_kwargs: object) -> Any:
            token.cancel("test")
            return _ocr_success("PAGE 1")

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", cancel_after_first)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_four_frame_tiff_path),
                staging,
                "md",
                {"to_md_enable_ocr": True, "to_md_keep_images": True},
                source_format="tif",
            )
            context._cancellation = token
            with pytest.raises(CancellationRequested):
                ImageToMarkdownConverter().convert(context)
            assert list(Path(staging).iterdir()) == []
            assert context.workspace.registered_artifacts == []

    @pytest.mark.contract
    def test_tiff_to_markdown_frame_materialization_failure_leaves_no_partial_outputs(
        self,
        sample_four_frame_tiff_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        original_save = converter_mod.save_image_with_options
        calls = 0

        def fail_third(*args: Any, **kwargs: Any) -> None:
            nonlocal calls
            calls += 1
            if calls == 3:
                raise OSError("third frame cannot be written")
            original_save(*args, **kwargs)

        monkeypatch.setattr(converter_mod, "save_image_with_options", fail_third)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_four_frame_tiff_path),
                staging,
                "md",
                {"to_md_enable_ocr": False, "to_md_keep_images": True},
                source_format="tif",
            )
            result = ImageToMarkdownConverter().convert(context)
            assert list(Path(staging).iterdir()) == []

        assert result.success is False
        assert result.artifacts == []
        assert context.workspace.registered_artifacts == []

    @pytest.mark.integration
    def test_tiff_runtime_merges_one_bound_ocr_warning_per_fragment(
        self,
        sample_four_frame_tiff_path: Path,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Runtime progress/result merging must not duplicate per-frame warnings."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_core.text.ocr import OcrOutcome, OcrStatus
        from docwen_plugin_image.plugin import ImagePlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        monkeypatch.setattr(
            converter_mod,
            "run_ocr_outcome",
            lambda *_args, **_kwargs: OcrOutcome(OcrStatus.NO_TEXT),
        )
        registry = PluginRegistry()
        registry.register(ImagePlugin())
        task_manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "workspace")),
            OutputFinalizer(),
        )
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        request = ConversionRequest(
            request_id="tiff-physical-page-warning-merge",
            input_refs=[
                FileRef(
                    path=str(sample_four_frame_tiff_path),
                    format="tif",
                    category="image",
                    size_bytes=sample_four_frame_tiff_path.stat().st_size,
                )
            ],
            target_format="md",
            options={"to_md_enable_ocr": True, "to_md_keep_images": False},
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_manager.execute_single(request)

        assert result.success is True
        fragments = [artifact for artifact in result.artifacts if artifact.kind == "auxiliary"]
        warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.code == "OCR-BEST-EFFORT"]
        assert len(fragments) == 4
        assert len(warnings) == 4
        assert [diagnostic.artifact_id for diagnostic in warnings] == [artifact.artifact_id for artifact in fragments]
