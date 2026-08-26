"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

import pytest

from ._layout_conversions_support import (
    Path,
    _build_fake_context,
    _create_interactive_geometry_pdf,
    _create_metadata_pdf,
    _create_pdf_ops_fixture_inputs,
    _create_rotated_pdf,
    _create_text_pdf,
    _expected_interactive_geometry_projection,
    _load_pdf_ops_old_system_fixture,
    _pdf_geometry_projection,
    _pdf_interactive_geometry_projection,
    _pdf_metadata_projection,
    _pdf_page_texts,
    tempfile,
)

pytestmark = pytest.mark.contract


class TestPluginDispatch:
    def test_plugin_dispatches_png(self, sample_pdf_path: Path) -> None:
        """LayoutPlugin should dispatch PDF→PNG correctly."""
        from docwen_plugin_layout import LayoutPlugin

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "png",
            )
            result = LayoutPlugin().convert(context)
            assert result.success is True
            assert len(result.artifacts) == 1
            assert result.artifacts[0].media_type == "image/png"

    def test_plugin_dispatches_pdf_to_md(self, sample_pdf_path: Path) -> None:
        """LayoutPlugin should dispatch pdf→md correctly."""
        from docwen_plugin_layout import LayoutPlugin

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "md",
                options={"to_md_keep_images": False},
            )
            result = LayoutPlugin().convert(context)
            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) >= 1
            assert result.artifacts[0].media_type == "text/markdown"

    def test_plugin_dispatches_merge_pdfs_action(self, tmp_path: Path) -> None:
        """Named merge_pdfs actions must merge all PDF input refs in old-system order."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        expected = fixture["projects"]["docwen-current"]["merge"]
        paths = _create_pdf_ops_fixture_inputs(tmp_path)
        first = paths["first.pdf"]
        second = paths["second.pdf"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(first),
                staging,
                "pdf",
                action_name="merge_pdfs",
            )
            context.request.input_refs.append(
                type(context.request.input_refs[0])(
                    path=str(second),
                    format="pdf",
                    category="layout",
                )
            )

            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.suggested_name == expected["suggested_name"]
            assert artifact.media_type == expected["artifact_media_type"]
            assert artifact.metadata["input_count"] == expected["metadata_input_count"]
            assert result.metrics.extra["input_count"] == expected["metrics_input_count"]
            assert _pdf_page_texts(artifact.staging_path) == expected["page_texts"]

    def test_merge_and_split_open_admitted_pdf_content_with_misleading_suffixes(self, tmp_path: Path) -> None:
        """Both PDF operations keep the admitted parser across arbitrary names."""
        from docwen_plugin_layout import LayoutPlugin

        first = tmp_path / "first.txt"
        second = tmp_path / "second.data"
        split_source = tmp_path / "pages.markdown"
        _create_text_pdf(first, ["FIRST CONTENT"])
        _create_text_pdf(second, ["SECOND CONTENT"])
        _create_text_pdf(split_source, ["PAGE ONE", "PAGE TWO", "PAGE THREE"])

        merge_staging = tmp_path / "merge-staging"
        merge_staging.mkdir()
        merge_context = _build_fake_context(
            str(first),
            str(merge_staging),
            "pdf",
            action_name="merge_pdfs",
            source_format="pdf",
        )
        merge_context.request.input_refs.append(
            type(merge_context.request.input_refs[0])(
                path=str(second),
                format="pdf",
                category="layout",
            )
        )
        merge_result = LayoutPlugin().convert(merge_context)

        split_staging = tmp_path / "split-staging"
        split_staging.mkdir()
        split_context = _build_fake_context(
            str(split_source),
            str(split_staging),
            "pdf",
            options={"split_mode": "every_page"},
            action_name="split_pdf",
            source_format="pdf",
        )
        split_result = LayoutPlugin().convert(split_context)

        assert merge_result.success is True, f"unexpected merge error: {merge_result.error}"
        assert _pdf_page_texts(merge_result.artifacts[0].staging_path) == ["FIRST CONTENT", "SECOND CONTENT"]
        assert split_result.success is True, f"unexpected split error: {split_result.error}"
        assert [_pdf_page_texts(artifact.staging_path) for artifact in split_result.artifacts] == [
            ["PAGE ONE"],
            ["PAGE TWO"],
            ["PAGE THREE"],
        ]

    def test_merge_pdfs_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path) -> None:
        """PDF merge should finalize the merged PDF into the user output dir."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_layout import LayoutPlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture = _load_pdf_ops_old_system_fixture()
        expected = fixture["projects"]["docwen-current"]["merge"]
        paths = _create_pdf_ops_fixture_inputs(tmp_path)
        merge_inputs = [paths["first.pdf"], paths["second.pdf"]]
        output_dir = tmp_path / "out-merge"
        output_dir.mkdir()
        workspace_root = tmp_path / "workspace-merge"
        registry = PluginRegistry()
        registry.register(LayoutPlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id="pdf-merge-finalizer-old-system-fixture",
            input_refs=[
                FileRef(
                    path=str(path),
                    format="pdf",
                    category="layout",
                    size_bytes=path.stat().st_size,
                )
                for path in merge_inputs
            ],
            target_format="pdf",
            action_name="merge_pdfs",
            options={},
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_mgr.execute_single(request)

        assert result.success, f"unexpected error: {result.error}"
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        output_path = Path(artifact.staging_path)
        assert output_path.parent == output_dir
        assert output_path.name == expected["suggested_name"]
        assert output_path.suffix.lower() == expected["output_suffix"]
        assert artifact.media_type == expected["artifact_media_type"]
        assert artifact.metadata["input_count"] == expected["metadata_input_count"]
        assert output_path.stat().st_size > 0
        assert _pdf_page_texts(output_path) == expected["page_texts"]
        assert any(d.code == "PDF-MERGE-OK" for d in result.diagnostics)
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
        assert str(workspace_root) not in str(output_path)

    def test_merge_pdfs_reports_corrupt_input_file_name(self, sample_pdf_path: Path, tmp_path: Path) -> None:
        """Corrupt PDFs should identify the offending file in user-facing errors."""
        from docwen_plugin_layout import LayoutPlugin

        broken = tmp_path / "broken.pdf"
        broken.write_bytes(b"%PDF-1.4\nnot a real pdf body\n")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "pdf",
                action_name="merge_pdfs",
            )
            context.request.input_refs.append(
                type(context.request.input_refs[0])(
                    path=str(broken),
                    format="pdf",
                    category="layout",
                )
            )

            result = LayoutPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"
            assert "broken.pdf" in result.error.message
            assert "Failed to merge" in result.error.message

    def test_pdf_operations_metadata_boundary_matches_old_system_projection(self, tmp_path: Path) -> None:
        """PDF merge/split preserve pages but clear document metadata like the old systems."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        probe = fixture["metadata_boundary_probe"]
        expected_empty_metadata = probe["expected_empty_metadata"]
        pdf_inputs = probe["input_pdfs"]
        paths: dict[str, Path] = {}

        for item in [*pdf_inputs["merge"], pdf_inputs["split"]]:
            path = tmp_path / item["name"]
            _create_metadata_pdf(
                path,
                {
                    "page_size": [300, 180],
                    "page_texts": item["page_texts"],
                    "metadata": item["metadata"],
                },
            )
            paths[item["name"]] = path
            input_projection = _pdf_metadata_projection(path)
            assert input_projection["page_texts"] == item["page_texts"]
            assert input_projection["metadata"] == item["metadata"]

        with tempfile.TemporaryDirectory() as staging:
            merge_first, merge_second = pdf_inputs["merge"]
            context = _build_fake_context(
                str(paths[merge_first["name"]]),
                staging,
                "pdf",
                action_name="merge_pdfs",
            )
            context.request.input_refs.append(
                type(context.request.input_refs[0])(
                    path=str(paths[merge_second["name"]]),
                    format="pdf",
                    category="layout",
                )
            )

            merge_result = LayoutPlugin().convert(context)

            assert merge_result.success is True, f"unexpected error: {merge_result.error}"
            assert len(merge_result.artifacts) == 1
            merge_projection = _pdf_metadata_projection(merge_result.artifacts[0].staging_path)
            assert merge_projection["pdf_magic"] == "%PDF-"
            assert merge_projection["page_count"] == 2
            assert merge_projection["page_texts"] == probe["projects"]["docwen-current"]["merge_page_texts"]
            assert merge_projection["metadata"] == expected_empty_metadata
            assert merge_result.artifacts[0].metadata["input_count"] == 2

        with tempfile.TemporaryDirectory() as staging:
            split_input = pdf_inputs["split"]
            context = _build_fake_context(
                str(paths[split_input["name"]]),
                staging,
                "pdf",
                options={"split_mode": "custom", "pages": [1, 3]},
                action_name="split_pdf",
            )

            split_result = LayoutPlugin().convert(context)

            assert split_result.success is True, f"unexpected error: {split_result.error}"
            assert len(split_result.artifacts) == 2
            projections = [_pdf_metadata_projection(artifact.staging_path) for artifact in split_result.artifacts]
            assert [projection["pdf_magic"] for projection in projections] == ["%PDF-", "%PDF-"]
            assert [projection["page_count"] for projection in projections] == [2, 2]
            assert [projection["page_texts"] for projection in projections] == probe["projects"]["docwen-current"][
                "split_custom_artifact_page_groups"
            ]
            assert [projection["metadata"] for projection in projections] == [
                expected_empty_metadata,
                expected_empty_metadata,
            ]
            assert [artifact.metadata["pages"] for artifact in split_result.artifacts] == [[1, 3], [2, 4]]

    def test_pdf_operations_rotated_page_boundary_matches_old_system_projection(self, tmp_path: Path) -> None:
        """PDF merge/split preserve rotated page geometry like the old systems."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        probe = fixture["rotated_page_boundary_probe"]
        paths: dict[str, Path] = {}

        for item in probe["input_pdfs"]["merge"]:
            path = tmp_path / item["name"]
            _create_rotated_pdf(path, item)
            paths[item["name"]] = path
            assert _pdf_geometry_projection(path) == item["expected_source_projection"]

        with tempfile.TemporaryDirectory() as staging:
            merge_first, merge_second = probe["input_pdfs"]["merge"]
            context = _build_fake_context(
                str(paths[merge_first["name"]]),
                staging,
                "pdf",
                action_name="merge_pdfs",
            )
            context.request.input_refs.append(
                type(context.request.input_refs[0])(
                    path=str(paths[merge_second["name"]]),
                    format="pdf",
                    category="layout",
                )
            )

            merge_result = LayoutPlugin().convert(context)

            assert merge_result.success is True, f"unexpected error: {merge_result.error}"
            assert len(merge_result.artifacts) == 1
            merge_projection = _pdf_geometry_projection(merge_result.artifacts[0].staging_path)
            assert merge_projection == probe["projects"]["docwen-current"]["merge_page_projection"]

        with tempfile.TemporaryDirectory() as staging:
            split_input = probe["input_pdfs"]["split"]
            split_path = tmp_path / split_input["name"]
            _create_rotated_pdf(split_path, split_input)
            assert _pdf_geometry_projection(split_path) == split_input["expected_source_projection"]
            context = _build_fake_context(
                str(split_path),
                staging,
                "pdf",
                options={"split_mode": "custom", "pages": [1]},
                action_name="split_pdf",
            )

            split_result = LayoutPlugin().convert(context)

            assert split_result.success is True, f"unexpected error: {split_result.error}"
            assert len(split_result.artifacts) == 2
            projections = [_pdf_geometry_projection(artifact.staging_path) for artifact in split_result.artifacts]
            assert projections == probe["projects"]["docwen-current"]["split_custom_artifact_projections"]

    def test_pdf_operations_interactive_geometry_matches_old_system_projection(self, tmp_path: Path) -> None:
        """Merge/split preserve heterogeneous geometry, URI links, and text notes."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        probe = fixture["interactive_geometry_boundary_probe"]
        page_profiles = probe["page_profiles"]
        inputs = probe["input_pdfs"]
        paths: dict[str, Path] = {}

        for item in [*inputs["merge"], inputs["split"]]:
            path = tmp_path / item["name"]
            _create_interactive_geometry_pdf(path, item, page_profiles)
            paths[item["name"]] = path

        with tempfile.TemporaryDirectory() as staging:
            merge_first, merge_second = inputs["merge"]
            context = _build_fake_context(
                str(paths[merge_first["name"]]),
                staging,
                "pdf",
                action_name="merge_pdfs",
            )
            context.request.input_refs.append(
                type(context.request.input_refs[0])(
                    path=str(paths[merge_second["name"]]),
                    format="pdf",
                    category="layout",
                )
            )

            merge_result = LayoutPlugin().convert(context)

            assert merge_result.success is True, f"unexpected error: {merge_result.error}"
            assert len(merge_result.artifacts) == 1
            expected_merge = _expected_interactive_geometry_projection(
                [*merge_first["pages"], *merge_second["pages"]],
                page_profiles,
            )
            assert _pdf_interactive_geometry_projection(merge_result.artifacts[0].staging_path) == expected_merge

        with tempfile.TemporaryDirectory() as staging:
            split_input = inputs["split"]
            context = _build_fake_context(
                str(paths[split_input["name"]]),
                staging,
                "pdf",
                options={"split_mode": "custom", "pages": [1, 3]},
                action_name="split_pdf",
            )

            split_result = LayoutPlugin().convert(context)

            assert split_result.success is True, f"unexpected error: {split_result.error}"
            assert len(split_result.artifacts) == 2
            expected_groups = [
                [split_input["pages"][0], split_input["pages"][2]],
                [split_input["pages"][1], split_input["pages"][3]],
            ]
            assert [
                _pdf_interactive_geometry_projection(artifact.staging_path) for artifact in split_result.artifacts
            ] == [_expected_interactive_geometry_projection(group, page_profiles) for group in expected_groups]
            assert [artifact.metadata["pages"] for artifact in split_result.artifacts] == [[1, 3], [2, 4]]
