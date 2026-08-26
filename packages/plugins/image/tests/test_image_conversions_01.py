"""Focused tests split from test_image_conversions.py."""

from __future__ import annotations

from ._image_conversions_support import (
    PROJECT_ROOT,
    Any,
    Image,
    Path,
    _build_fake_context,
    hashlib,
    json,
    os,
    pytest,
    struct,
    tempfile,
)

pytestmark = pytest.mark.golden


class TestImageFormatConversion:
    @pytest.mark.contract
    def test_png_to_jpg_preserves_size_and_sets_media_type(self, sample_png_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_png_path), staging, "jpg")
            result = ImageFormatConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.media_type == "image/jpeg"
            assert artifact.suggested_name == "sample.jpg"
            assert os.path.isfile(artifact.staging_path)

            with Image.open(artifact.staging_path) as img:
                assert img.format == "JPEG"
                assert img.size == (32, 24)
                assert img.mode == "RGB"

            # L-3: metrics/diagnostics assertions
            assert result.metrics.input_bytes > 0
            assert result.metrics.output_bytes > 0
            assert result.metrics.extra["target_format"] == "jpg"
            assert any(d.code == "IMAGEFMT-OK" for d in result.diagnostics)

    @pytest.mark.contract
    def test_png_to_image_formats_match_old_system_semantic_fixture(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        expected = fixture["projects"]["docwen-current"]
        sample_png_path = tmp_path / fixture["input_image"]["name"]
        img = Image.new("RGBA", tuple(fixture["input_image"]["size"]), (20, 120, 200, fixture["input_image"]["alpha"]))
        img.save(sample_png_path, format="PNG")
        img.close()

        for target in fixture["targets"]:
            with tempfile.TemporaryDirectory() as staging:
                context = _build_fake_context(
                    str(sample_png_path),
                    staging,
                    target,
                    {"compress_mode": "lossless"},
                )
                result = ImageFormatConverter().convert(context)

                assert result.success is True
                artifact = result.artifacts[0]
                expected_target = expected[target]

                assert artifact.media_type == expected_target["artifact_media_type"]
                assert artifact.suggested_name == expected_target["suggested_name"]
                assert artifact.metadata["target_format"] == expected_target["metadata_target_format"]
                assert result.metrics.extra["target_format"] == expected_target["metrics_target_format"]
                assert Path(artifact.staging_path).suffix.lower() == expected_target["suffix"]
                assert Path(artifact.staging_path).stat().st_size > 0

                with Image.open(artifact.staging_path) as img:
                    assert img.format == expected_target["format"]
                    assert img.mode == expected_target["mode"]
                    assert list(img.size) == expected_target["size"]
                    assert getattr(img, "n_frames", 1) == expected_target["n_frames"]

    @pytest.mark.contract
    def test_grayscale_palette_mode_matrix_matches_recorded_projection(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        probe = fixture["grayscale_palette_mode_probe"]
        size = tuple(probe["input_size"])
        input_paths: dict[str, Path] = {}

        gray = Image.new("L", size)
        gray.putdata([(x * 5 + y * 7) % 256 for y in range(size[1]) for x in range(size[0])])
        input_paths["gray_l"] = tmp_path / "gray_l.png"
        gray.save(input_paths["gray_l"], format="PNG")
        gray.close()

        palette = []
        for index in range(256):
            palette.extend(((index * 3) % 256, (index * 5) % 256, (index * 7) % 256))
        for input_name, transparent in (("palette_opaque", False), ("palette_transparent", True)):
            image = Image.new("P", size)
            image.putpalette(palette)
            image.putdata([(x + y * 3) % 16 for y in range(size[1]) for x in range(size[0])])
            input_paths[input_name] = tmp_path / f"{input_name}.png"
            save_kwargs: dict[str, Any] = {"format": "PNG"}
            if transparent:
                save_kwargs["transparency"] = 0
            image.save(input_paths[input_name], **save_kwargs)
            image.close()

        expected_current = probe["projects"]["docwen-current"]
        for input_name, input_path in input_paths.items():
            for target in probe["targets"]:
                scenario = f"{input_name}_to_{target}"
                expected = expected_current[scenario]
                with tempfile.TemporaryDirectory() as staging:
                    context = _build_fake_context(
                        str(input_path),
                        staging,
                        target,
                        {"compress_mode": "lossless"},
                    )
                    result = ImageFormatConverter().convert(context)

                    assert result.success is True
                    assert len(result.artifacts) == 1
                    artifact = result.artifacts[0]
                    assert artifact.metadata["target_format"] == target
                    assert artifact.suggested_name == f"{input_name}.{target}"
                    assert Path(artifact.staging_path).stat().st_size > 0
                    with Image.open(artifact.staging_path) as converted:
                        converted.load()
                        assert converted.format == expected["format"]
                        assert converted.mode == expected["mode"]
                        assert list(converted.size) == probe["shared_output_size"]
                        assert getattr(converted, "n_frames", 1) == probe["shared_n_frames"]
                        rgba = converted.convert("RGBA")
                        try:
                            assert list(rgba.getchannel("A").getextrema()) == expected["alpha_extrema"]
                            if target == "png":
                                assert hashlib.sha256(rgba.tobytes()).hexdigest() == expected["rgba_sha256"]
                        finally:
                            rgba.close()

    @pytest.mark.contract
    def test_extended_native_mode_matrix_matches_recorded_projection(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_extended_mode_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        size = tuple(fixture["input_size"])
        input_paths: dict[str, Path] = {}

        cmyk = Image.new("CMYK", size)
        cmyk.putdata(
            [
                ((x * 5) % 256, (y * 9) % 256, ((x + y) * 7) % 256, ((x * 3 + y * 2) % 80))
                for y in range(size[1])
                for x in range(size[0])
            ]
        )
        input_paths["cmyk"] = tmp_path / fixture["inputs"]["cmyk"]["name"]
        cmyk.save(input_paths["cmyk"], format="TIFF", compression="tiff_lzw")
        cmyk.close()

        binary = Image.new("1", size)
        binary.putdata([(x // 4 + y // 4) % 2 for y in range(size[1]) for x in range(size[0])])
        input_paths["binary_1"] = tmp_path / fixture["inputs"]["binary_1"]["name"]
        binary.save(input_paths["binary_1"], format="PNG")
        binary.close()

        values16 = [(x * 1301 + y * 2089) % 65536 for y in range(size[1]) for x in range(size[0])]
        raw16 = struct.pack(f"<{len(values16)}H", *values16)
        gray16 = Image.frombytes("I;16", size, raw16)
        input_paths["gray16"] = tmp_path / fixture["inputs"]["gray16"]["name"]
        gray16.save(input_paths["gray16"], format="PNG")
        gray16.close()

        float_image = Image.new("F", size)
        float_image.putdata([((x * 1.75) + (y * 2.5)) % 255.0 for y in range(size[1]) for x in range(size[0])])
        input_paths["float_f"] = tmp_path / fixture["inputs"]["float_f"]["name"]
        float_image.save(input_paths["float_f"], format="TIFF")
        float_image.close()

        media_types = fixture["current_artifact_contract"]["media_types"]
        for input_name, input_path in input_paths.items():
            for target in fixture["targets"]:
                scenario = fixture["scenarios"][f"{input_name}_to_{target}"]
                expected = scenario["current_projection"]
                with tempfile.TemporaryDirectory() as staging:
                    context = _build_fake_context(
                        str(input_path),
                        staging,
                        target,
                        {"compress_mode": "lossless"},
                    )
                    result = ImageFormatConverter().convert(context)

                    assert result.success is True
                    assert len(result.artifacts) == 1
                    artifact = result.artifacts[0]
                    assert artifact.media_type == media_types[target]
                    assert artifact.suggested_name == f"{input_path.stem}.{target}"
                    assert artifact.metadata["target_format"] == target
                    assert Path(artifact.staging_path).stat().st_size > 0
                    with Image.open(artifact.staging_path) as converted:
                        converted.load()
                        assert converted.format == expected["format"]
                        assert converted.mode == expected["mode"]
                        assert list(converted.size) == fixture["input_size"]
                        assert getattr(converted, "n_frames", 1) == fixture["shared_n_frames"]
                        assert json.loads(json.dumps(converted.getextrema())) == expected["extrema"]
                        rgba = converted.convert("RGBA")
                        try:
                            if target != "webp":
                                assert hashlib.sha256(rgba.tobytes()).hexdigest() == expected["rgba_sha256"]
                        finally:
                            rgba.close()

    @pytest.mark.contract
    def test_exif_orientation_probe_matches_old_system_raw_pixel_boundary(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        probe = fixture["exif_orientation_probe"]
        input_spec = probe["input_image"]
        input_path = tmp_path / input_spec["name"]

        img = Image.new(input_spec["mode"], tuple(input_spec["size"]), "white")
        for x in range(8):
            for y in range(8):
                img.putpixel((x, y), (255, 0, 0))
        for x in range(input_spec["size"][0] - 8, input_spec["size"][0]):
            for y in range(input_spec["size"][1] - 8, input_spec["size"][1]):
                img.putpixel((x, y), (0, 0, 255))
        exif = Image.Exif()
        exif[274] = input_spec["exif_orientation"]
        img.save(input_path, input_spec["format"], exif=exif)
        img.close()

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(input_path), staging, probe["target"])
            result = ImageFormatConverter().convert(context)

            assert result.success is True
            artifact = result.artifacts[0]
            expected = probe["expected_projection"]
            assert artifact.metadata["width"] == expected["size"][0]
            assert artifact.metadata["height"] == expected["size"][1]
            with Image.open(artifact.staging_path) as converted:
                assert converted.format == expected["format"]
                assert converted.mode == expected["mode"]
                assert list(converted.size) == expected["size"]
                assert converted.getexif().get(274) == expected["exif_orientation"]

    @pytest.mark.contract
    def test_exif_metadata_probe_matches_old_system_clearing_boundary(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        # The selector name is retained as historical fixture provenance. VIS-201
        # intentionally improves current JPG export over both clearing references.
        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        probe = fixture["exif_metadata_probe"]
        input_spec = probe["input_image"]
        input_path = tmp_path / input_spec["name"]

        img = Image.new(input_spec["mode"], tuple(input_spec["size"]), (80, 120, 160))
        exif = Image.Exif()
        exif[271] = input_spec["exif"]["make"]
        exif[272] = input_spec["exif"]["model"]
        exif[274] = input_spec["exif"]["orientation"]
        exif[306] = input_spec["exif"]["datetime"]
        exif[315] = input_spec["exif"]["artist"]
        img.save(input_path, input_spec["format"], exif=exif)
        img.close()

        with Image.open(input_path) as source:
            source_exif = source.getexif()
            assert source_exif.get(271) == input_spec["exif"]["make"]
            assert source_exif.get(272) == input_spec["exif"]["model"]
            assert source_exif.get(274) == input_spec["exif"]["orientation"]
            assert source_exif.get(306) == input_spec["exif"]["datetime"]
            assert source_exif.get(315) == input_spec["exif"]["artist"]
            assert len(source_exif) == input_spec["exif_tag_count"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(input_path), staging, probe["target"])
            result = ImageFormatConverter().convert(context)

            assert result.success is True
            artifact = result.artifacts[0]
            expected_current = probe["projects"]["docwen-current"]
            assert artifact.media_type == expected_current["artifact_media_type"]
            assert artifact.suggested_name == expected_current["suggested_name"]
            assert artifact.metadata["target_format"] == expected_current["metadata_target_format"]
            assert artifact.metadata["width"] == input_spec["size"][0]
            assert artifact.metadata["height"] == input_spec["size"][1]
            with Image.open(artifact.staging_path) as converted:
                converted_exif = converted.getexif()
                assert converted.format == probe["expected_projection"]["format"]
                assert converted.mode == probe["expected_projection"]["mode"]
                assert list(converted.size) == input_spec["size"]
                assert converted_exif.get(271) == input_spec["exif"]["make"]
                assert converted_exif.get(272) == input_spec["exif"]["model"]
                assert converted_exif.get(274) == 1
                assert converted_exif.get(306) == input_spec["exif"]["datetime"]
                assert converted_exif.get(315) is None
                assert len(converted_exif) == 4

    @pytest.mark.parametrize("target", ["jpg", "webp"])
    @pytest.mark.contract
    def test_fa08_flat_export_normalizes_orientation_and_preserves_supported_metadata(
        self, tmp_path: Path, target: str
    ) -> None:
        from PIL import ImageChops, ImageOps, ImageStat

        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        input_path = tmp_path / "fa08-orientation-metadata.jpg"
        source = Image.new("RGB", (40, 20), "white")
        for x in range(40):
            for y in range(20):
                source.putpixel((x, y), (220, 20, 20) if x < 20 else (20, 80, 220))
        exif = Image.Exif()
        exif[271] = "DocWenMake"
        exif[272] = "DocWenModel"
        exif[274] = 6
        exif[305] = "DocWenSoftware"
        exif[306] = "2026:07:23 12:34:56"
        exif[315] = "unsupported-artist"
        icc_profile = bytes(range(256)) * 2 + bytes(range(24))
        source.save(input_path, format="JPEG", quality=100, exif=exif, icc_profile=icc_profile)
        source.close()
        source_sha = hashlib.sha256(input_path.read_bytes()).hexdigest()

        with Image.open(input_path) as persisted_source:
            expected_display = ImageOps.exif_transpose(persisted_source).convert("RGB")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(input_path),
                staging,
                target,
                {"compress_mode": "lossless"},
            )
            result = ImageFormatConverter().convert(context)

            assert result.success is True
            assert hashlib.sha256(input_path.read_bytes()).hexdigest() == source_sha
            artifact = result.artifacts[0]
            assert artifact.metadata["width"] == 20
            assert artifact.metadata["height"] == 40
            with Image.open(artifact.staging_path) as converted:
                converted.load()
                assert converted.size == (20, 40)
                assert converted.info.get("icc_profile") == icc_profile
                converted_exif = converted.getexif()
                assert converted_exif.get(271) == "DocWenMake"
                assert converted_exif.get(272) == "DocWenModel"
                assert converted_exif.get(274) == 1
                assert converted_exif.get(305) == "DocWenSoftware"
                assert converted_exif.get(306) == "2026:07:23 12:34:56"
                assert converted_exif.get(315) is None
                assert converted_exif.get(34853) is None
                displayed_again = ImageOps.exif_transpose(converted).convert("RGB")
                actual_display = converted.convert("RGB")
                try:
                    assert displayed_again.size == converted.size
                    difference = ImageStat.Stat(ImageChops.difference(expected_display, actual_display))
                    similarity = 1.0 - (sum(difference.mean) / 3.0 / 255.0)
                    assert similarity >= 0.95
                finally:
                    displayed_again.close()
                    actual_display.close()
                    expected_display.close()

    @pytest.mark.contract
    def test_expanded_bmp_gif_tif_matrix_matches_old_system_projection(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        probe = fixture["expanded_format_matrix_probe"]
        input_spec = probe["input_image"]
        input_path = tmp_path / input_spec["name"]

        img = Image.new(input_spec["mode"], tuple(input_spec["size"]), "white")
        for x in range(input_spec["size"][0]):
            for y in range(input_spec["size"][1]):
                color = (220, 30, 40) if x < 10 else ((30, 120, 220) if y < 12 else (40, 180, 90))
                img.putpixel((x, y), color)
        img.save(input_path, input_spec["format"])
        img.close()

        expected = probe["projects"]["docwen-current"]
        for target in probe["targets"]:
            with tempfile.TemporaryDirectory() as staging:
                context = _build_fake_context(
                    str(input_path),
                    staging,
                    target,
                    {"compress_mode": "lossless"},
                )
                result = ImageFormatConverter().convert(context)

                assert result.success is True
                artifact = result.artifacts[0]
                expected_target = expected[target]
                assert artifact.media_type == expected_target["artifact_media_type"]
                assert artifact.suggested_name == expected_target["suggested_name"]
                assert artifact.metadata["target_format"] == expected_target["metadata_target_format"]
                assert result.metrics.extra["target_format"] == expected_target["metrics_target_format"]
                assert Path(artifact.staging_path).suffix.lower() == expected_target["suffix"]
                assert Path(artifact.staging_path).stat().st_size > 0
                with Image.open(artifact.staging_path) as converted:
                    assert converted.format == expected_target["format"]
                    assert converted.mode == expected_target["mode"]
                    assert list(converted.size) == expected_target["size"]
                    assert getattr(converted, "n_frames", 1) == expected_target["n_frames"]

    @pytest.mark.contract
    def test_multiframe_input_boundary_matches_recorded_projection(self, tmp_path: Path) -> None:
        from docwen_plugin_image.format_conversion.converter import ImageFormatConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        probe = fixture["multiframe_input_boundary_probe"]
        gif_path = tmp_path / "animated_rgb.gif"
        tiff_path = tmp_path / "multi_rgb.tif"

        gif_frames = [
            Image.new("RGB", tuple(probe["inputs"]["animated_rgb.gif"]["size"]), color)
            for color in (
                (220, 20, 20),
                (20, 120, 220),
                (40, 180, 80),
            )
        ]
        gif_frames[0].save(gif_path, format="GIF", save_all=True, append_images=gif_frames[1:], duration=80, loop=0)
        for frame in gif_frames:
            frame.close()

        tiff_frames = [
            Image.new("RGB", tuple(probe["inputs"]["multi_rgb.tif"]["size"]), color)
            for color in (
                (180, 30, 30),
                (30, 150, 90),
                (40, 80, 200),
            )
        ]
        tiff_frames[0].save(
            tiff_path,
            format="TIFF",
            save_all=True,
            append_images=tiff_frames[1:],
            compression="tiff_lzw",
        )
        for frame in tiff_frames:
            frame.close()

        input_paths = {
            "animated_rgb.gif": gif_path,
            "multi_rgb.tif": tiff_path,
        }
        expected_project = probe["projects"]["docwen-current"]
        for scenario_id, scenario in probe["scenarios"].items():
            with tempfile.TemporaryDirectory() as staging:
                context = _build_fake_context(
                    str(input_paths[scenario["input"]]),
                    staging,
                    scenario["target"],
                    {"compress_mode": "lossless"},
                )
                result = ImageFormatConverter().convert(context)

                expected = expected_project[scenario_id]
                assert result.success is True
                assert len(result.artifacts) == expected["artifact_count"]
                assert result.metrics.extra["frame_count"] == expected["metrics_frame_count"]
                for artifact, expected_artifact in zip(result.artifacts, expected["artifacts"], strict=True):
                    artifact_path = Path(artifact.staging_path)
                    assert artifact.kind == expected_artifact["kind"]
                    assert artifact.is_primary is expected_artifact["is_primary"]
                    assert artifact.suggested_name == expected_artifact["suggested_name"]
                    assert artifact.media_type == expected_artifact["media_type"]
                    assert artifact_path.suffix.lower() == expected_artifact["suffix"]
                    assert artifact_path.stat().st_size > 0
                    with Image.open(artifact_path) as converted:
                        assert converted.format == expected_artifact["format"]
                        assert converted.mode == expected_artifact["mode"]
                        assert list(converted.size) == expected_artifact["size"]
                        assert getattr(converted, "n_frames", 1) == expected_artifact["n_frames"]
