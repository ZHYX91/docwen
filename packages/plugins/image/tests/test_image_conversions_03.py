"""Focused tests split from test_image_conversions.py."""

from __future__ import annotations

from ._image_conversions_support import (
    PROJECT_ROOT,
    Any,
    Image,
    Path,
    _build_fake_context,
    _deliverable_artifacts,
    _document_node_root,
    _ocr_success,
    json,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.golden


class TestImageToMarkdown:
    @pytest.mark.contract
    def test_image_to_markdown_matches_old_system_semantic_fixture(self, tmp_path: Path) -> None:
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_to_markdown_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        expected = fixture["projects"]["docwen-current"]

        sample_png_path = tmp_path / fixture["input_image"]["name"]
        Image.new("RGB", tuple(fixture["input_image"]["size"]), (20, 120, 200)).save(sample_png_path, format="PNG")

        config_snapshot = {
            "link": {
                "format": {
                    "image_link_style": "markdown_embed",
                    "md_file_link_style": "markdown_link",
                }
            },
            "conversion": {"ocr_output": {"show_blockquote_title": False}},
            "export": {"to_md_ocr_placement_mode": "main_md"},
        }
        scenarios = {
            "file_no_ocr": {"image_mode": "file", "to_md_keep_images": True, "to_md_enable_ocr": False},
            "base64_no_ocr": {"image_mode": "base64", "to_md_keep_images": True, "to_md_enable_ocr": False},
        }
        for scenario_id in fixture["scenarios"]:
            with tempfile.TemporaryDirectory() as staging:
                context = _build_fake_context(
                    str(sample_png_path),
                    staging,
                    "md",
                    scenarios[scenario_id],
                    config_snapshot=config_snapshot,
                )
                result = ImageToMarkdownConverter().convert(context)

                assert result.success is True
                md_artifact = result.artifacts[0]
                md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
                image_artifacts = [artifact for artifact in result.artifacts if artifact.kind == "image"]
                expected_scenario = expected[scenario_id]

                assert md_artifact.media_type == expected_scenario["primary_media_type"]
                assert md_artifact.suggested_name == expected_scenario["primary_suggested_name"]
                assert md_artifact.metadata["image_mode"] == expected_scenario["metadata_image_mode"]
                assert md_artifact.metadata["keep_images"] == expected_scenario["metadata_keep_images"]
                assert md_artifact.metadata["ocr_enabled"] == expected_scenario["metadata_ocr_enabled"]
                assert result.metrics.extra["artifact_count"] == expected_scenario["metrics_artifact_count"]
                assert result.metrics.extra["ocr_enabled"] == expected_scenario["metrics_ocr_enabled"]
                assert Path(md_artifact.staging_path).suffix.lower() == expected_scenario["output_suffix"]
                assert (
                    len([artifact for artifact in result.artifacts if artifact.media_type == "text/markdown"])
                    == expected_scenario["md_file_count"]
                )
                assert len(image_artifacts) == expected_scenario["png_file_count"]
                assert [artifact.suggested_name for artifact in image_artifacts] == expected_scenario[
                    "image_suggested_names"
                ]
                assert md_text.startswith("---") is expected_scenario["has_yaml_frontmatter"]
                assert ("source_format: png" in md_text) is expected_scenario["contains_source_format"]
                assert ("data:image/" in md_text) is expected_scenario["contains_data_uri"]
                assert ("![" in md_text) is expected_scenario["contains_markdown_image_link"]
                assert ("![[" in md_text) is expected_scenario["contains_wiki_image_link"]
                assert ("sample_rgb" in md_text) is expected_scenario["contains_sample_png"]

    @pytest.mark.contract
    def test_image_to_markdown_file_mode_registers_md_and_image_artifacts(self, sample_png_path: Path) -> None:
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_keep_images": True, "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 2
            md_artifact = result.artifacts[0]
            image_artifact = result.artifacts[1]
            assert md_artifact.media_type == "text/markdown"
            assert image_artifact.kind == "image"
            md_text = Path(md_artifact.staging_path).read_text(encoding="utf-8")
            assert "source_format: png" in md_text
            assert "![[sample.png]]" in md_text
            # L-3: metrics
            assert result.metrics.input_bytes > 0
            assert result.metrics.output_bytes > 0
            assert result.metrics.extra["artifact_count"] == 2

    @pytest.mark.integration
    def test_image_to_markdown_file_mode_artifacts_are_finalized_through_runtime(
        self,
        sample_png_path: Path,
        tmp_path: Path,
    ) -> None:
        """Image→Markdown file mode must place Markdown and image artifacts."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_image.plugin import ImagePlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        output_dir = tmp_path / "out"
        output_dir.mkdir()
        registry = PluginRegistry()
        registry.register(ImagePlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "workspace")),
            OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id="image-to-md-finalizer-old-system-fixture",
            input_refs=[FileRef(path=str(sample_png_path), format="png", category="image")],
            target_format="md",
            action_name="",
            options={
                "image_mode": "file",
                "to_md_keep_images": True,
                "to_md_enable_ocr": False,
                "yaml_key_labels": {"title": "Titel"},
            },
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_mgr.execute_single(request)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        assert len(_deliverable_artifacts(result)) == 2
        md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
        image_artifact = next(artifact for artifact in result.artifacts if artifact.kind == "image")
        md_path = Path(md_artifact.staging_path)
        image_path = Path(image_artifact.staging_path)
        node_root = _document_node_root(md_path, output_dir)
        assert _document_node_root(image_path, output_dir) == node_root
        assert md_path.name == f"{node_root.name}.md"
        assert md_artifact.metadata["source_suggested_name"] == "sample.md"
        assert image_path.name == "sample.png"
        md_text = md_path.read_text(encoding="utf-8")
        assert "Titel: sample" in md_text
        assert "title: sample" not in md_text
        assert "![[sample.png]]" in md_text
        assert str(tmp_path / "workspace") not in md_text
        assert image_path.read_bytes() == sample_png_path.read_bytes()

    @pytest.mark.integration
    def test_image_to_markdown_same_dir_reuses_retained_input_without_orphan_collision(
        self,
        sample_png_path: Path,
        tmp_path: Path,
    ) -> None:
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_image.plugin import ImagePlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        registry = PluginRegistry()
        registry.register(ImagePlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "workspace-same-dir")),
            OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id="image-to-md-same-dir-retained-input",
            input_refs=[FileRef(path=str(sample_png_path), format="png", category="image")],
            target_format="md",
            action_name="",
            options={
                "image_mode": "file",
                "to_md_keep_images": True,
                "to_md_enable_ocr": False,
            },
            output_policy=OutputPolicy(overwrite_mode="rename"),
        )

        result = task_mgr.execute_single(request)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        md_artifact = next(artifact for artifact in result.artifacts if artifact.media_type == "text/markdown")
        image_artifact = next(artifact for artifact in result.artifacts if artifact.kind == "image")
        md_path = Path(md_artifact.staging_path)
        image_path = Path(image_artifact.staging_path)
        node_root = _document_node_root(md_path, sample_png_path.parent)
        assert _document_node_root(image_path, sample_png_path.parent) == node_root
        assert image_path.name == sample_png_path.name
        assert sample_png_path.exists()
        assert image_path.read_bytes() == sample_png_path.read_bytes()
        assert not sample_png_path.with_name(f"{sample_png_path.stem}_001.png").exists()
        assert f"![[{sample_png_path.name}]]" in md_path.read_text(encoding="utf-8")

    @pytest.mark.integration
    def test_image_to_markdown_base64_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path) -> None:
        """Image->Markdown base64 mode should finalize only the primary Markdown."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_image.plugin import ImagePlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_to_markdown_semantics.json"
        fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
        expected = fixture["projects"]["docwen-current"]["base64_no_ocr"]
        input_path = tmp_path / fixture["input_image"]["name"]
        Image.new("RGB", tuple(fixture["input_image"]["size"]), (20, 120, 200)).save(input_path, format="PNG")
        output_dir = tmp_path / "out"
        output_dir.mkdir()
        workspace_root = tmp_path / "workspace"
        registry = PluginRegistry()
        registry.register(ImagePlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id="image-to-md-base64-finalizer-old-system-fixture",
            input_refs=[FileRef(path=str(input_path), format="png", category="image")],
            target_format="md",
            action_name="",
            options={"image_mode": "base64", "to_md_keep_images": True, "to_md_enable_ocr": False},
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_mgr.execute_single(request)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        assert len(_deliverable_artifacts(result)) == expected["metrics_artifact_count"]
        artifact = _deliverable_artifacts(result)[0]
        md_path = Path(artifact.staging_path)
        md_text = md_path.read_text(encoding="utf-8")
        node_root = _document_node_root(md_path, output_dir)
        assert sorted(path.name for path in output_dir.iterdir()) == [node_root.name]
        assert md_path.name == f"{node_root.name}.md"
        assert artifact.metadata["source_suggested_name"] == expected["primary_suggested_name"]
        assert artifact.media_type == expected["primary_media_type"]
        assert artifact.metadata["image_mode"] == expected["metadata_image_mode"]
        assert artifact.metadata["keep_images"] == expected["metadata_keep_images"]
        assert artifact.metadata["ocr_enabled"] == expected["metadata_ocr_enabled"]
        assert ("data:image/" in md_text) is expected["contains_data_uri"]
        assert len([item for item in result.artifacts if item.kind == "image"]) == expected["png_file_count"]
        assert str(workspace_root) not in md_text
        assert str(output_dir) not in md_text
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

    @pytest.mark.integration
    def test_image_to_markdown_image_md_ocr_sidecar_finalizes_through_runtime(
        self,
        sample_png_path: Path,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Image->Markdown image_md OCR mode must finalize primary, image, and sidecar artifacts."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_image.plugin import ImagePlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        monkeypatch.setattr(
            converter_mod,
            "run_ocr_outcome",
            lambda _path, **_kwargs: _ocr_success("OCR sidecar text"),
        )

        output_dir = tmp_path / "out"
        output_dir.mkdir()
        workspace_root = tmp_path / "workspace"
        registry = PluginRegistry()
        registry.register(ImagePlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id="image-to-md-ocr-sidecar-finalizer",
            input_refs=[FileRef(path=str(sample_png_path), format="png", category="image")],
            target_format="md",
            action_name="",
            options={
                "image_mode": "file",
                "to_md_keep_images": True,
                "to_md_enable_ocr": True,
                "ocr_placement": "image_md",
                "yaml_key_labels": {"title": "Titel"},
            },
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_mgr.execute_single(request)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        assert len(_deliverable_artifacts(result)) == 3
        primary = next(artifact for artifact in result.artifacts if artifact.kind == "primary")
        image = next(artifact for artifact in result.artifacts if artifact.kind == "image")
        sidecar = next(artifact for artifact in result.artifacts if artifact.kind == "auxiliary")
        primary_path = Path(primary.staging_path)
        image_path = Path(image.staging_path)
        sidecar_path = Path(sidecar.staging_path)

        node_root = _document_node_root(primary_path, output_dir)
        assert _document_node_root(image_path, output_dir) == node_root
        assert _document_node_root(sidecar_path, output_dir) == node_root
        assert primary_path.name == f"{node_root.name}.md"
        assert primary.metadata["source_suggested_name"] == "sample.md"
        assert image_path.name == "sample.png"
        assert sidecar.metadata["source_suggested_name"] == "sample_ocr.md"
        assert sidecar.metadata["ocr"] is True
        assert sidecar.media_type == "text/markdown"

        primary_text = primary_path.read_text(encoding="utf-8")
        sidecar_text = sidecar_path.read_text(encoding="utf-8")
        assert "Titel: sample" in primary_text
        assert sidecar.logical_path is not None
        assert f"![[{sidecar.logical_path.split('/', 1)[1]}]]" in primary_text
        assert "OCR sidecar text" not in primary_text
        assert "Titel: sample_ocr" in sidecar_text
        assert "![[../sample.png]]" in sidecar_text
        assert "> OCR sidecar text" in sidecar_text
        assert image_path.read_bytes() == sample_png_path.read_bytes()
        assert str(workspace_root) not in primary_text
        assert str(workspace_root) not in sidecar_text
        assert any(d.code == "IMG2MD-OCR-OK" for d in result.diagnostics)
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

    @pytest.mark.contract
    def test_image_to_markdown_base64_embeds_data_uri(self, sample_png_path: Path) -> None:
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path), staging, "md", {"image_mode": "base64", "to_md_enable_ocr": False}
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "data:image/png;base64," in md_text
            assert result.metrics.output_bytes > 0

    @pytest.mark.contract
    def test_image_to_markdown_omit_mode_produces_html_comment(self, sample_png_path: Path) -> None:
        """L-6: image_mode=omit should produce an HTML comment placeholder."""
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path), staging, "md", {"image_mode": "omit", "to_md_enable_ocr": False}
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1  # Only the MD artifact
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "<!-- image omitted:" in md_text
            assert "source_format: png" in md_text
            assert result.metrics.input_bytes > 0

    @pytest.mark.contract
    def test_image_to_markdown_ocr_failure_is_best_effort(
        self,
        sample_png_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Optional OCR failure must preserve the base image conversion."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        def missing_ocr_model(_path: str, **_kwargs: object) -> Any:
            raise FileNotFoundError("rapidocr model missing")

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", missing_ocr_model)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path), staging, "md", {"to_md_enable_ocr": True, "to_md_keep_images": True}
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert result.error is None
            assert any(item[2] == "OCR-BEST-EFFORT" for item in context.progress.diagnostics)
            assert not any(d.code == "IMG2MD-OCR-OK" for d in result.diagnostics)
            assert Path(result.artifacts[0].staging_path).is_file()

    @pytest.mark.parametrize(
        "status_value",
        ["input_missing", "unavailable", "model_missing", "initialization_failed", "recognition_failed"],
    )
    @pytest.mark.contract
    def test_image_to_markdown_typed_ocr_failures_are_safe_and_nonfatal(
        self,
        sample_png_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        status_value: str,
    ) -> None:
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_core.text.ocr import OcrOutcome, OcrStatus
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        monkeypatch.setattr(
            converter_mod,
            "run_ocr_outcome",
            lambda *_args, **_kwargs: OcrOutcome(
                OcrStatus(status_value),
                message="private model path",
            ),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"to_md_enable_ocr": True, "to_md_keep_images": True},
            )
            result = ImageToMarkdownConverter().convert(context)

        assert result.success is True
        assert len(context.progress.diagnostics) == 1
        warning = context.progress.diagnostics[0]
        assert warning[2] == "OCR-BEST-EFFORT"
        assert f"status={status_value}" in warning[1]
        assert "private" not in warning[1]
        assert not any(d.code == "IMG2MD-OCR-OK" for d in result.diagnostics)

    @pytest.mark.contract
    def test_image_to_markdown_no_text_warns_about_possible_missed_text(
        self,
        sample_png_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_core.text.ocr import OcrOutcome, OcrStatus
        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        monkeypatch.setattr(
            converter_mod,
            "run_ocr_outcome",
            lambda *_args, **_kwargs: OcrOutcome(OcrStatus.NO_TEXT),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"to_md_enable_ocr": True, "to_md_keep_images": True},
            )
            result = ImageToMarkdownConverter().convert(context)

        assert result.success is True
        assert len(context.progress.diagnostics) == 1
        warning = context.progress.diagnostics[0]
        assert warning[2] == "OCR-BEST-EFFORT"
        assert "status=no_text" in warning[1]
        assert "may have been missed" in warning[1]
        assert not any(d.code == "IMG2MD-OCR-OK" for d in result.diagnostics)
