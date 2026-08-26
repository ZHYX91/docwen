"""Focused tests split from test_image_conversions.py."""

from __future__ import annotations

from ._image_conversions_support import (
    PROJECT_ROOT,
    BytesIO,
    Image,
    Path,
    PngImagePlugin,
    _assert_pdf_embedded_image_projection,
    _build_fake_context,
    _write_exif_jpeg,
    hashlib,
    json,
    os,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.golden


class TestImageFormatConversion:
    @pytest.mark.integration
    def test_image_format_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path) -> None:
        """Image format conversion should be finalized into the user output dir."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_image.plugin import ImagePlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        input_path = tmp_path / fixture["input_image"]["name"]
        Image.new(
            fixture["input_image"]["mode"],
            tuple(fixture["input_image"]["size"]),
            (20, 120, 200, fixture["input_image"]["alpha"]),
        ).save(input_path, format="PNG")
        output_dir = tmp_path / "out"
        output_dir.mkdir()
        registry = PluginRegistry()
        registry.register(ImagePlugin())
        workspace_root = tmp_path / "workspace"
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )
        for target in fixture["targets"]:
            target_output_dir = output_dir / target
            target_output_dir.mkdir()
            request = ConversionRequest(
                request_id=f"image-format-finalizer-old-system-fixture-{target}",
                input_refs=[
                    FileRef(
                        path=str(input_path),
                        format="png",
                        category="image",
                        size_bytes=input_path.stat().st_size,
                    )
                ],
                target_format=target,
                action_name="",
                options={"compress_mode": "lossless"},
                output_policy=OutputPolicy(output_dir=str(target_output_dir)),
            )

            result = task_mgr.execute_single(request)

            expected = fixture["projects"]["docwen-current"][target]
            assert result.success, f"unexpected error for {target}: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            image_path = Path(artifact.staging_path)
            assert image_path.parent == target_output_dir
            assert image_path.name == expected["suggested_name"]
            assert artifact.media_type == expected["artifact_media_type"]
            assert artifact.metadata["target_format"] == expected["metadata_target_format"]
            assert image_path.suffix.lower() == expected["suffix"]
            assert image_path.stat().st_size > 0
            assert any(d.code == "IMAGEFMT-OK" for d in result.diagnostics)
            assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
            with Image.open(image_path) as img:
                assert img.format == expected["format"]
                assert img.mode == expected["mode"]
                assert list(img.size) == expected["size"]
                assert getattr(img, "n_frames", 1) == expected["n_frames"]
            assert str(workspace_root) not in str(image_path)

    @pytest.mark.contract
    def test_jpg_to_png_cross_format_preserves_dimensions(self, sample_jpg_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_jpg_path), staging, "png")
            result = ImageFormatConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.media_type == "image/png"
            with Image.open(artifact.staging_path) as img:
                assert img.format == "PNG"
                assert img.size == (48, 36)
            assert result.metrics.input_bytes > 0
            assert result.metrics.output_bytes > 0

    @pytest.mark.contract
    def test_multipage_tiff_to_png_creates_one_artifact_per_frame(self, sample_tiff_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_tiff_path), staging, "png")
            result = ImageFormatConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 2
            assert result.artifacts[0].is_primary is True
            assert result.artifacts[1].kind == "auxiliary"
            for artifact in result.artifacts:
                with Image.open(artifact.staging_path) as img:
                    assert img.format == "PNG"
                    assert img.size == (12, 10)
            assert result.metrics.extra["frame_count"] == 2

    @pytest.mark.contract
    def test_limit_size_for_large_image_fails_when_target_too_small(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        large_path = tmp_path / "large.png"
        raw = os.urandom(200 * 200 * 3)
        img = Image.frombytes("RGB", (200, 200), raw)
        img.save(large_path, format="PNG")
        img.close()

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(large_path), staging, "png", {"compress_mode": "limit_size", "size_limit": 1, "size_unit": "KB"}
            )
            result = ImageFormatConverter().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.diagnostic_code == "IMAGEFMT-CONVERT-ERROR"
            assert any(d.code == "IMAGEFMT-CONVERT-ERROR" for d in result.diagnostics)

    @pytest.mark.contract
    def test_compression_limit_probe_matches_old_system_projection(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
        probe = json.loads(fixture_path.read_text(encoding="utf-8"))["compression_limit_probe"]
        input_spec = probe["input_image"]
        input_path = tmp_path / input_spec["name"]
        img = Image.new(input_spec["mode"], tuple(input_spec["size"]))
        width, height = input_spec["size"]
        pixels = [
            ((x * 37 + y * 17) % 256, (x * 13 + y * 57) % 256, (x * 91 + y * 23) % 256)
            for y in range(height)
            for x in range(width)
        ]
        img.putdata(pixels)
        base_png = BytesIO()
        img.save(base_png, format=input_spec["format"], compress_level=7)
        # zlib output differs across supported operating systems. Preserve the
        # frozen byte-size oracle without changing pixels by adding one
        # deterministic, ignored PNG text chunk sized from the local encoder.
        text_chunk_overhead = 12 + len("pad") + 1
        padding_size = input_spec["source_bytes"] - len(base_png.getvalue()) - text_chunk_overhead
        assert padding_size >= 0, "local PNG encoder exceeded the frozen source-size oracle"
        png_info = PngImagePlugin.PngInfo()
        png_info.add_text("pad", "x" * padding_size, zip=False)
        img.save(input_path, format=input_spec["format"], compress_level=7, pnginfo=png_info)
        img.close()

        assert input_path.stat().st_size == input_spec["source_bytes"]

        expected_project = probe["projects"]["docwen-current"]
        for scenario in probe["scenarios"]:
            with tempfile.TemporaryDirectory() as staging:
                context = _build_fake_context(
                    str(input_path),
                    staging,
                    scenario["target"],
                    scenario["options"],
                )
                result = ImageFormatConverter().convert(context)

                expected = expected_project[scenario["id"]]
                assert result.success is expected["success"]
                if expected["success"]:
                    assert len(result.artifacts) == 1
                    artifact = result.artifacts[0]
                    artifact_path = Path(artifact.staging_path)
                    assert artifact.media_type == expected["artifact_media_type"]
                    assert artifact.suggested_name == expected["suggested_name"]
                    assert artifact.metadata["target_format"] == expected["metadata_target_format"]
                    assert artifact_path.stat().st_size <= scenario["limit_bytes"]
                    with Image.open(artifact_path) as converted:
                        assert converted.format == expected["format"]
                        assert converted.mode == expected["mode"]
                        assert list(converted.size) == expected["size"]
                else:
                    assert result.error is not None
                    assert result.error.diagnostic_code == expected["diagnostic_code"]
                    assert len(result.artifacts) == expected["artifact_count"]
                    assert any(d.code == expected["diagnostic_code"] for d in result.diagnostics)

    @pytest.mark.contract
    def test_bmp_to_jpg_converts_transparency_to_rgb(self, sample_bmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_bmp_path), staging, "jpg")
            result = ImageFormatConverter().convert(context)

            assert result.success is True
            artifact = result.artifacts[0]
            with Image.open(artifact.staging_path) as img:
                assert img.format == "JPEG"
                assert img.mode == "RGB"
                assert img.size == (24, 18)


class TestImageToPdf:
    @pytest.mark.contract
    def test_fa08_mpo_delivers_every_frame_with_auxiliary_warning(self, tmp_path: Path) -> None:
        fitz = pytest.importorskip("fitz")
        from docwen_plugin_image.to_pdf.converter import ImageToPdfConverter

        input_path = tmp_path / "two-frame-phone.jpg"
        primary = Image.new("RGB", (24, 16), (220, 40, 40))
        auxiliary = Image.new("RGB", (24, 16), (90, 90, 90))
        primary.save(input_path, format="MPO", save_all=True, append_images=[auxiliary])
        primary.close()
        auxiliary.close()
        source_sha = hashlib.sha256(input_path.read_bytes()).hexdigest()

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(input_path), staging, "pdf", {"quality_mode": "original"})
            result = ImageToPdfConverter().convert(context)

            assert result.success is True
            assert hashlib.sha256(input_path.read_bytes()).hexdigest() == source_sha
            warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.level == "warning"]
            assert len(warnings) == 1
            assert warnings[0].code == "IMG2PDF-MPO-AUXILIARY-FRAMES"
            assert "2" in warnings[0].message
            assert "auxiliary" in warnings[0].message.lower()
            with fitz.open(result.artifacts[0].staging_path) as document:
                assert len(document) == 2

    @pytest.mark.contract
    def test_fa08_single_frame_pdf_does_not_emit_mpo_warning(self, tmp_path: Path) -> None:
        from docwen_plugin_image.to_pdf.converter import ImageToPdfConverter

        input_path = tmp_path / "single-frame.jpg"
        Image.new("RGB", (24, 16), (220, 40, 40)).save(input_path, format="JPEG")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(input_path), staging, "pdf", {"quality_mode": "original"})
            result = ImageToPdfConverter().convert(context)

            assert result.success is True
            assert not [
                diagnostic for diagnostic in result.diagnostics if diagnostic.code == "IMG2PDF-MPO-AUXILIARY-FRAMES"
            ]

    @pytest.mark.contract
    def test_image_to_pdf_matches_old_system_semantic_fixture(self, tmp_path: Path) -> None:
        fitz = pytest.importorskip("fitz")
        from docwen_plugin_image.to_pdf.converter import ImageToPdfConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_to_pdf_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        expected = fixture["projects"]["docwen-current"]

        png_path = tmp_path / "sample_rgb.png"
        Image.new("RGB", tuple(fixture["inputs"]["sample_rgb.png"]["size"]), (20, 120, 200)).save(
            png_path, format="PNG"
        )

        tiff_path = tmp_path / "multi_rgb.tif"
        tiff_input = fixture["inputs"]["multi_rgb.tif"]
        frames = [
            Image.new("RGB", tuple(tiff_input["frame_size"]), color)
            for color in ((255, 255, 255), (128, 128, 128), (0, 0, 0))
        ]
        frames[0].save(tiff_path, save_all=True, append_images=frames[1:], format="TIFF", compression="tiff_lzw")
        for frame in frames:
            frame.close()

        exif_path = tmp_path / "exif_pdf_probe.jpg"
        _write_exif_jpeg(exif_path, fixture["inputs"]["exif_pdf_probe.jpg"])

        scenarios = {
            "png_original": (png_path, "original"),
            "png_a4": (png_path, "a4"),
            "multipage_tiff_original": (tiff_path, "original"),
            "jpeg_exif_original": (exif_path, "original"),
        }
        for scenario_id in fixture["scenarios"]:
            input_path, quality_mode = scenarios[scenario_id]
            with tempfile.TemporaryDirectory() as staging:
                context = _build_fake_context(
                    str(input_path),
                    staging,
                    "pdf",
                    {"quality_mode": quality_mode},
                )
                result = ImageToPdfConverter().convert(context)

                assert result.success is True
                artifact = result.artifacts[0]
                expected_scenario = expected[scenario_id]
                artifact_path = Path(artifact.staging_path)

                assert artifact.media_type == expected_scenario["artifact_media_type"]
                assert artifact.suggested_name == expected_scenario["suggested_name"]
                assert artifact.metadata["quality_mode"] == expected_scenario["metadata_quality_mode"]
                assert result.metrics.extra["quality_mode"] == expected_scenario["metrics_quality_mode"]
                assert artifact_path.suffix.lower() == expected_scenario["output_suffix"]
                assert artifact_path.stat().st_size > 0
                assert artifact_path.read_bytes().startswith(b"%PDF") is expected_scenario["pdf_magic"]

                with fitz.open(artifact_path) as doc:
                    assert doc.page_count == expected_scenario["page_count"]
                    rects = [[round(value, 2) for value in doc[index].rect] for index in range(doc.page_count)]
                    assert rects == expected_scenario["page_rects"]
                    if "pdf_document_metadata" in expected_scenario:
                        assert {
                            key: doc.metadata.get(key)
                            for key in ("title", "author", "subject", "keywords", "creator", "producer")
                        } == expected_scenario["pdf_document_metadata"]
                    if "embedded_image_projection" in expected_scenario:
                        _assert_pdf_embedded_image_projection(doc, expected_scenario["embedded_image_projection"])

    @pytest.mark.integration
    def test_image_to_pdf_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path) -> None:
        """Image→PDF old-system fixture scenarios should finalize into the user output dir."""
        fitz = pytest.importorskip("fitz")
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_image.plugin import ImagePlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_to_pdf_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        expected = fixture["projects"]["docwen-current"]
        png_path = tmp_path / "sample_rgb.png"
        Image.new("RGB", tuple(fixture["inputs"]["sample_rgb.png"]["size"]), (20, 120, 200)).save(
            png_path, format="PNG"
        )

        tiff_path = tmp_path / "multi_rgb.tif"
        tiff_input = fixture["inputs"]["multi_rgb.tif"]
        frames = [
            Image.new("RGB", tuple(tiff_input["frame_size"]), color)
            for color in ((255, 255, 255), (128, 128, 128), (0, 0, 0))
        ]
        frames[0].save(tiff_path, save_all=True, append_images=frames[1:], format="TIFF", compression="tiff_lzw")
        for frame in frames:
            frame.close()

        exif_path = tmp_path / "exif_pdf_probe.jpg"
        _write_exif_jpeg(exif_path, fixture["inputs"]["exif_pdf_probe.jpg"])

        registry = PluginRegistry()
        registry.register(ImagePlugin())
        workspace_root = tmp_path / "workspace"
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )
        scenarios = {
            "png_original": (png_path, "original"),
            "png_a4": (png_path, "a4"),
            "multipage_tiff_original": (tiff_path, "original"),
            "jpeg_exif_original": (exif_path, "original"),
        }

        for scenario_id in fixture["scenarios"]:
            input_path, quality_mode = scenarios[scenario_id]
            output_dir = tmp_path / f"out_{scenario_id}"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id=f"image-to-pdf-finalizer-{scenario_id}",
                input_refs=[
                    FileRef(
                        path=str(input_path),
                        format={
                            "png_original": "png",
                            "png_a4": "png",
                            "multipage_tiff_original": "tiff",
                            "jpeg_exif_original": "jpeg",
                        }[scenario_id],
                        category="image",
                        size_bytes=input_path.stat().st_size,
                    )
                ],
                target_format="pdf",
                action_name="",
                options={"quality_mode": quality_mode},
                output_policy=OutputPolicy(output_dir=str(output_dir)),
            )

            result = task_mgr.execute_single(request)

            assert result.success, f"unexpected error for {scenario_id}: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            expected_scenario = expected[scenario_id]
            pdf_path = Path(artifact.staging_path)
            assert pdf_path.parent == output_dir
            assert pdf_path.name == expected_scenario["suggested_name"]
            assert artifact.media_type == expected_scenario["artifact_media_type"]
            assert artifact.metadata["quality_mode"] == expected_scenario["metadata_quality_mode"]
            assert pdf_path.suffix.lower() == expected_scenario["output_suffix"]
            assert pdf_path.stat().st_size > 0
            assert pdf_path.read_bytes().startswith(b"%PDF") is expected_scenario["pdf_magic"]
            assert any(d.code == "IMG2PDF-OK" for d in result.diagnostics)
            assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
            with fitz.open(pdf_path) as doc:
                assert doc.page_count == expected_scenario["page_count"]
                rects = [[round(value, 2) for value in doc[index].rect] for index in range(doc.page_count)]
                assert rects == expected_scenario["page_rects"]
                if "pdf_document_metadata" in expected_scenario:
                    assert {
                        key: doc.metadata.get(key)
                        for key in ("title", "author", "subject", "keywords", "creator", "producer")
                    } == expected_scenario["pdf_document_metadata"]
                if "embedded_image_projection" in expected_scenario:
                    _assert_pdf_embedded_image_projection(doc, expected_scenario["embedded_image_projection"])
            assert str(workspace_root) not in str(pdf_path)

    @pytest.mark.contract
    def test_png_to_pdf_creates_pdf_artifact(self, sample_png_path: Path) -> None:
        from docwen_plugin_image.to_pdf.converter import ImageToPdfConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_png_path), staging, "pdf", {"quality_mode": "original"})
            result = ImageToPdfConverter().convert(context)

            assert result.success is True
            artifact = result.artifacts[0]
            assert artifact.media_type == "application/pdf"
            assert artifact.suggested_name == "sample.pdf"
            assert Path(artifact.staging_path).read_bytes().startswith(b"%PDF")
            # L-3: metrics
            assert result.metrics.input_bytes > 0
            assert result.metrics.output_bytes > 0

    @pytest.mark.contract
    def test_png_to_pdf_with_explicit_quality_mode(self, sample_png_path: Path) -> None:
        """The canonical quality_mode option should work."""
        from docwen_plugin_image.to_pdf.converter import ImageToPdfConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_png_path), staging, "pdf", {"quality_mode": "original"})
            result = ImageToPdfConverter().convert(context)

            assert result.success is True
            assert Path(result.artifacts[0].staging_path).read_bytes().startswith(b"%PDF")
