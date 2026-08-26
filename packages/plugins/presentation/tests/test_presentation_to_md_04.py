"""Focused tests split from test_presentation_to_md.py."""

from __future__ import annotations

from ._presentation_to_md_support import (
    Path,
    _build_policy03_chart_pptx,
    _build_policy03_media_pptx,
    _document_node_root,
    _markdown_targets,
    _run_request,
    pytest,
    re,
)
from ._presentation_to_md_support import (
    pipeline as pipeline,
)

pytestmark = [pytest.mark.golden, pytest.mark.contract]
from ._presentation_to_md_support import (
    sample_pptx_file as sample_pptx_file,
)
from ._presentation_to_md_support import (
    sample_pptx_with_image as sample_pptx_with_image,
)
from ._presentation_to_md_support import (
    sample_pptx_with_notes as sample_pptx_with_notes,
)
from ._presentation_to_md_support import (
    sample_pptx_with_sections as sample_pptx_with_sections,
)


class TestPptToMd:
    """ROUTE-PPT-001 uses preprocessing bridge before PPTX → MD."""

    def test_ppt_to_md_rejects_empty_input(self, pipeline, tmp_path) -> None:
        """Empty PPT input should fail fast before invoking the bridge."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_ppt"
        output_dir.mkdir()

        dummy_ppt = tmp_path / "dummy.ppt"
        dummy_ppt.write_bytes(b"")

        result = _run_request(task_mgr, dummy_ppt, "ppt", output_dir)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "invalid_input"
        assert "ppt" in result.error.message.lower()

    def test_ppt_to_md_keeps_original_input_stem_after_hub_preprocessing(
        self,
        pipeline,
        sample_pptx_file,
        tmp_path,
        monkeypatch,
    ) -> None:
        """The intermediate PPTX path must not leak into final names or fallback title."""
        from pptx import Presentation

        from docwen_core.office_bridge import BridgeResult

        _plugin, task_mgr, _ws_mgr = pipeline
        source = tmp_path / "legacy-presentation.ppt"
        source.write_bytes(b"legacy ppt placeholder")
        hub_pptx = tmp_path / "auxiliary_1.pptx"
        presentation = Presentation(sample_pptx_file)
        presentation.core_properties.title = ""
        presentation.save(hub_pptx)

        monkeypatch.setattr(
            "docwen_plugin_presentation.pptx_md.ppt_converter.convert_with_backend_priority",
            lambda *_args, **_kwargs: BridgeResult(
                True,
                output_path=str(hub_pptx),
                backend="test PowerPoint",
            ),
        )
        output_dir = tmp_path / "output_ppt"
        output_dir.mkdir()

        result = _run_request(task_mgr, source, "ppt", output_dir)

        assert result.success is True
        primary = next(artifact for artifact in result.artifacts if artifact.kind == "primary")
        assert primary.metadata["source_suggested_name"] == "legacy-presentation.md"
        node_root = _document_node_root(Path(primary.staging_path), output_dir)
        assert Path(primary.staging_path).name == f"{node_root.name}.md"
        content = Path(primary.staging_path).read_text(encoding="utf-8")
        assert "title: legacy-presentation" in content
        assert "# legacy-presentation" in content
        assert "auxiliary_1" not in content

    def test_ppt_to_md_finalizes_images_registered_through_hub_workspace(
        self,
        pipeline,
        sample_pptx_with_image,
        tmp_path,
        monkeypatch,
    ) -> None:
        """PPT bridge delegation must not drop downstream image artifacts."""
        from docwen_core.office_bridge import BridgeResult

        _plugin, task_mgr, ws_mgr = pipeline
        source = tmp_path / "legacy-with-image.ppt"
        source.write_bytes(b"legacy ppt placeholder")
        monkeypatch.setattr(
            "docwen_plugin_presentation.pptx_md.ppt_converter.convert_with_backend_priority",
            lambda *_args, **_kwargs: BridgeResult(
                True,
                output_path=str(sample_pptx_with_image),
                backend="test PowerPoint",
            ),
        )
        output_dir = tmp_path / "output_ppt_with_image"
        output_dir.mkdir()

        result = _run_request(task_mgr, source, "ppt", output_dir)

        assert result.success is True
        primary = next(artifact for artifact in result.artifacts if artifact.kind == "primary")
        image = next(artifact for artifact in result.artifacts if artifact.kind == "image")
        assert primary.metadata["source_suggested_name"] == "legacy-with-image.md"
        assert re.fullmatch(
            r"legacy-with-image_slide1_img1_[0-9a-f]{12}\.png",
            image.suggested_name,
        )
        node_root = _document_node_root(Path(primary.staging_path), output_dir)
        assert _document_node_root(Path(image.staging_path), output_dir) == node_root
        assert (
            Path(image.staging_path).read_bytes()
            == Path(sample_pptx_with_image).parent.joinpath("tiny.png").read_bytes()
        )
        content = Path(primary.staging_path).read_text(encoding="utf-8")
        assert f"![[{image.suggested_name}]]" in content
        assert str(Path(ws_mgr.root_dir)) not in content

    def test_ppt_bridge_preserves_request_owned_ocr_title(
        self,
        pipeline,
        sample_pptx_with_image,
        tmp_path,
        monkeypatch,
    ) -> None:
        """The PPT hub proxy must not discard TaskManager's frozen OCR title."""
        from docwen_core.office_bridge import BridgeResult
        from docwen_core.text.ocr import OcrOutcome, OcrStatus

        source = tmp_path / "legacy-request-policy.ppt"
        source.write_bytes(b"legacy ppt placeholder")
        monkeypatch.setattr(
            "docwen_plugin_presentation.pptx_md.ppt_converter.convert_with_backend_priority",
            lambda *_args, **_kwargs: BridgeResult(
                True,
                output_path=str(sample_pptx_with_image),
                backend="test PowerPoint",
            ),
        )
        monkeypatch.setattr(
            "docwen_core.text.ocr.run_ocr_outcome",
            lambda _path, **_kwargs: OcrOutcome(OcrStatus.SUCCESS, text="legacy OCR body"),
        )
        monkeypatch.setattr(
            "docwen_runtime.config.build_ocr_blockquote_title",
            lambda *_args, **_kwargs: "Request localized title",
        )
        output_dir = tmp_path / "output_ppt_request_policy"
        output_dir.mkdir()
        result = _run_request(
            pipeline[1],
            source,
            "ppt",
            output_dir,
            config_snapshot={
                "conversion": {"ocr_output": {"show_blockquote_title": True}},
                "export": {
                    "to_md_image_extraction_mode": "file",
                    "to_md_ocr_placement_mode": "main_md",
                },
                "gui": {"language": {"locale": "en_US"}},
            },
            to_md_enable_ocr=True,
        )

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "> **Request localized title**" in content
        assert "Global OCR title" not in content

    def test_ppt_to_md_honors_configured_presentation_backend_priority(
        self,
        pipeline,
        sample_pptx_file,
        tmp_path,
        monkeypatch,
    ) -> None:
        """PPT preprocessing must consume the persisted presentation backend order."""
        from docwen_core.office_bridge import BridgeResult

        _plugin, task_mgr, _ws_mgr = pipeline
        source = tmp_path / "configured-priority.ppt"
        source.write_bytes(b"legacy ppt placeholder")
        observed_priority: list[str] = []
        observed_candidates: set[str] = set()
        observed_source_formats: list[str] = []

        def fake_bridge(*_args, **kwargs):
            observed_source_formats.append(kwargs["source_format"])
            observed_priority.extend(kwargs["backend_priority"])
            observed_candidates.update(kwargs["com_candidates"])
            return BridgeResult(True, output_path=str(sample_pptx_file), backend="test PowerPoint")

        monkeypatch.setattr(
            "docwen_plugin_presentation.pptx_md.ppt_converter.convert_with_backend_priority",
            fake_bridge,
        )
        output_dir = tmp_path / "output_configured_priority"
        output_dir.mkdir()

        result = _run_request(
            task_mgr,
            source,
            "ppt",
            output_dir,
            config_snapshot={
                "software": {
                    "default_priority": {
                        "presentation_processors": [
                            "msoffice_powerpoint",
                            "libreoffice",
                            "wps_presentation",
                        ]
                    }
                }
            },
        )

        assert result.success is True
        assert observed_source_formats == ["ppt"]
        assert observed_priority == [
            "msoffice_powerpoint",
            "libreoffice",
            "wps_presentation",
        ]
        assert observed_candidates == {"wps_presentation", "msoffice_powerpoint"}


class TestPluginDispatch:
    """Verify PresentationPlugin correctly dispatches to the right converters."""

    def test_can_handle_presentation_routes(self) -> None:
        """can_handle must return True for PPTX and PPT routes."""
        from docwen_plugin_presentation import PresentationPlugin

        plugin = PresentationPlugin()
        assert plugin.can_handle("pptx", "md") is True
        assert plugin.can_handle("ppt", "md") is True

    def test_can_handle_rejects_non_presentation_routes(self) -> None:
        """can_handle must reject routes belonging to other plugins."""
        from docwen_plugin_presentation import PresentationPlugin

        plugin = PresentationPlugin()
        assert plugin.can_handle("docx", "md") is False
        assert plugin.can_handle("html", "md") is False
        assert plugin.can_handle("pdf", "md") is False

    def test_convert_dispatch_pptx(self, pipeline, sample_pptx_file, tmp_path) -> None:
        """plugin.convert() must dispatch pptx→md to PptxToMarkdownConverter."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_dispatch_pptx"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_pptx_file, "pptx", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "Test Presentation" in content


class TestPolicy03PreservedPayloads:
    def test_chart_projects_ordered_semantics_and_exact_workbook(self, pipeline, tmp_path: Path) -> None:
        _plugin, task_mgr, _ws_mgr = pipeline
        source, expected_workbook = _build_policy03_chart_pptx(tmp_path)
        result = _run_request(task_mgr, source, "pptx", tmp_path / "chart-output")

        assert result.success is True
        markdown_artifact = next(artifact for artifact in result.artifacts if artifact.is_primary)
        markdown = Path(markdown_artifact.staging_path).read_text(encoding="utf-8")
        for token in ("Quarterly Sales", "Sales", "1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr", "8.2"):
            assert token in markdown
        workbook = next(
            artifact
            for artifact in result.artifacts
            if artifact.media_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        assert Path(workbook.staging_path).read_bytes() == expected_workbook
        assert workbook.suggested_name in _markdown_targets(markdown)
        warnings = [
            diagnostic for diagnostic in result.diagnostics if diagnostic.code == "PPTX-CHART-SNAPSHOT-UNAVAILABLE"
        ]
        assert len(warnings) == 1
        assert warnings[0].location == "slide 1: chart 1"

    @pytest.mark.parametrize(
        ("kind", "media_type"),
        (("audio", "audio/mpeg"), ("video", "video/mp4")),
    )
    def test_media_preserves_exact_playback_payload_and_poster(
        self,
        pipeline,
        tmp_path: Path,
        kind: str,
        media_type: str,
    ) -> None:
        _plugin, task_mgr, _ws_mgr = pipeline
        source, expected_media, expected_poster = _build_policy03_media_pptx(tmp_path, kind)
        result = _run_request(task_mgr, source, "pptx", tmp_path / f"{kind}-output")

        assert result.success is True
        markdown_artifact = next(artifact for artifact in result.artifacts if artifact.is_primary)
        markdown = Path(markdown_artifact.staging_path).read_text(encoding="utf-8")
        media = next(artifact for artifact in result.artifacts if artifact.media_type == media_type)
        poster = next(artifact for artifact in result.artifacts if artifact.media_type == "image/png")
        assert Path(media.staging_path).read_bytes() == expected_media
        assert Path(poster.staging_path).read_bytes() == expected_poster
        targets = _markdown_targets(markdown)
        assert media.suggested_name in targets
        assert poster.suggested_name in targets
