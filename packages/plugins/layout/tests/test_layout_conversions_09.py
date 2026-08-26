"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

import pytest

from ._layout_conversions_support import (
    Any,
    Path,
    _build_fake_context,
    _create_forms_actions_pdf,
    _create_pdf_ops_fixture_inputs,
    _load_pdf_ops_old_system_fixture,
    _pdf_forms_actions_projection,
    _pdf_page_texts,
    tempfile,
)

pytestmark = pytest.mark.contract


class TestPluginDispatch:
    def test_pdf_operations_forms_actions_match_old_system_projection(self, tmp_path: Path) -> None:
        """Split preserves safe internal actions when both pages stay together."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        probe = fixture["forms_actions_boundary_probe"]
        inputs = probe["input_pdfs"]
        contract = probe["page_contract"]
        paths: dict[str, Path] = {}

        def expected_page(page_id: str, *, target_index: int | None = None, target_id: str = "") -> dict[str, Any]:
            links = []
            if target_index is not None:
                links.append(
                    {
                        "kind": 1,
                        "target_page_index": target_index,
                        "target_text": f"FORM-{target_id}\nVALUE-{target_id}",
                    }
                )
            return {
                "text": f"FORM-{page_id}\nVALUE-{page_id}",
                "links": links,
                "annotations": [
                    {
                        "type": "FileAttachment",
                        "content": f"CONTENT-{page_id}",
                        "title": f"AUTHOR-{page_id}",
                        "rect": contract["attachment_rect"],
                        "filename": f"attachment-{page_id}-unicode.txt",
                        "payload_size": 14,
                        "payload_sha256": contract["payload_sha256_by_page"][page_id],
                    }
                ],
                "widgets": [
                    {
                        "field_name": f"field_{page_id}",
                        "field_label": f"LABEL-{page_id}",
                        "field_type": 7,
                        "field_type_string": "Text",
                        "field_value": f"VALUE-{page_id}",
                        "rect": contract["widget_rect"],
                    }
                ],
            }

        for item in [*inputs["merge"], inputs["split"]]:
            path = tmp_path / item["name"]
            _create_forms_actions_pdf(path, item)
            paths[item["name"]] = path

        merge_first, merge_second = inputs["merge"]
        assert _pdf_forms_actions_projection(paths[merge_first["name"]]) == [
            expected_page("A1", target_index=1, target_id="A2"),
            expected_page("A2"),
        ]
        assert _pdf_forms_actions_projection(paths[inputs["split"]["name"]])[0]["links"] == [
            {"kind": 1, "target_page_index": 2, "target_text": "FORM-S3\nVALUE-S3"}
        ]

        with tempfile.TemporaryDirectory() as staging:
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
            assert _pdf_forms_actions_projection(merge_result.artifacts[0].staging_path) == [
                expected_page("A1", target_index=1, target_id="A2"),
                expected_page("A2"),
                expected_page("B1"),
            ]

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
            assert [_pdf_forms_actions_projection(artifact.staging_path) for artifact in split_result.artifacts] == [
                [expected_page("S1", target_index=1, target_id="S3"), expected_page("S3")],
                [expected_page("S2", target_index=1, target_id="S4"), expected_page("S4")],
            ]
            assert [artifact.metadata["pages"] for artifact in split_result.artifacts] == [[1, 3], [2, 4]]

    def test_pdf_custom_split_does_not_crosslink_omitted_goto_target(self, tmp_path: Path) -> None:
        """A GOTO is omitted when its target lands in the other split artifact."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        split_input = fixture["forms_actions_boundary_probe"]["input_pdfs"]["split"]
        split_path = tmp_path / split_input["name"]
        _create_forms_actions_pdf(split_path, split_input)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(split_path),
                staging,
                "pdf",
                options={"split_mode": "custom", "pages": [1]},
                action_name="split_pdf",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            projections = [_pdf_forms_actions_projection(artifact.staging_path) for artifact in result.artifacts]
            assert projections[0][0]["links"] == []
            assert projections[1][0]["links"] == [
                {
                    "kind": 1,
                    "target_page_index": 2,
                    "target_text": "FORM-S4\nVALUE-S4",
                }
            ]

    def test_plugin_dispatches_split_pdf_action(self, tmp_path: Path) -> None:
        """Custom split should expose old-system selected/remaining page groups as artifacts."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        expected = fixture["projects"]["docwen-current"]["split_custom"]
        sample_pdf_path = _create_pdf_ops_fixture_inputs(tmp_path)["sample.pdf"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "pdf",
                options=expected["options"],
                action_name="split_pdf",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 2
            assert [artifact.suggested_name for artifact in result.artifacts] == expected["artifact_suggested_names"]
            assert {artifact.metadata.get("split_mode") for artifact in result.artifacts} == {"custom"}
            assert [artifact.metadata.get("pages") for artifact in result.artifacts] == expected[
                "artifact_metadata_pages"
            ]
            assert [_pdf_page_texts(artifact.staging_path) for artifact in result.artifacts] == expected[
                "artifact_page_groups"
            ]
            assert result.metrics is not None
            assert result.metrics.extra["split_mode"] == expected["metrics_split_mode"]
            assert result.metrics.extra["part1_pages"] == expected["metrics_part1_pages"]
            assert result.metrics.extra["part2_pages"] == expected["metrics_part2_pages"]

    def test_split_pdf_custom_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path) -> None:
        """PDF custom split should finalize all split artifacts into the user output dir."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_layout import LayoutPlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture = _load_pdf_ops_old_system_fixture()
        expected = fixture["projects"]["docwen-current"]["split_custom"]
        sample_pdf_path = _create_pdf_ops_fixture_inputs(tmp_path)["sample.pdf"]
        output_dir = tmp_path / "out-split"
        output_dir.mkdir()
        workspace_root = tmp_path / "workspace-split"
        registry = PluginRegistry()
        registry.register(LayoutPlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id="pdf-split-finalizer-old-system-fixture",
            input_refs=[
                FileRef(
                    path=str(sample_pdf_path),
                    format="pdf",
                    category="layout",
                    size_bytes=sample_pdf_path.stat().st_size,
                )
            ],
            target_format="pdf",
            action_name="split_pdf",
            options=expected["options"],
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_mgr.execute_single(request)

        assert result.success, f"unexpected error: {result.error}"
        assert len(result.artifacts) == 2
        assert [Path(artifact.staging_path).parent for artifact in result.artifacts] == [output_dir, output_dir]
        assert [Path(artifact.staging_path).name for artifact in result.artifacts] == expected[
            "artifact_suggested_names"
        ]
        assert [Path(artifact.staging_path).suffix.lower() for artifact in result.artifacts] == [".pdf", ".pdf"]
        assert [artifact.media_type for artifact in result.artifacts] == ["application/pdf", "application/pdf"]
        assert [artifact.metadata.get("pages") for artifact in result.artifacts] == expected["artifact_metadata_pages"]
        assert [artifact.metadata.get("split_mode") for artifact in result.artifacts] == ["custom", "custom"]
        assert [Path(artifact.staging_path).stat().st_size > 0 for artifact in result.artifacts] == [True, True]
        assert [_pdf_page_texts(artifact.staging_path) for artifact in result.artifacts] == expected[
            "artifact_page_groups"
        ]
        assert any(d.code == "PDF-SPLIT-OK" for d in result.diagnostics)
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
        assert all(str(workspace_root) not in str(artifact.staging_path) for artifact in result.artifacts)

    def test_split_pdf_every_page_and_odd_even_old_system_fixture_finalize_through_runtime(
        self, tmp_path: Path
    ) -> None:
        """PDF every-page and odd/even split should finalize all artifacts into the user output dir."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_layout import LayoutPlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture = _load_pdf_ops_old_system_fixture()
        current = fixture["projects"]["docwen-current"]
        sample_pdf_path = _create_pdf_ops_fixture_inputs(tmp_path)["sample.pdf"]
        workspace_root = tmp_path / "workspace-split-modes"
        registry = PluginRegistry()
        registry.register(LayoutPlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(workspace_root)),
            OutputFinalizer(),
        )
        scenarios = [
            ("every_page", current["split_every_page"]),
            ("odd_even", current["split_odd_even"]),
        ]

        for split_mode, expected in scenarios:
            output_dir = tmp_path / f"out-split-{split_mode}"
            output_dir.mkdir()
            request = ConversionRequest(
                request_id=f"pdf-split-finalizer-old-system-fixture-{split_mode}",
                input_refs=[
                    FileRef(
                        path=str(sample_pdf_path),
                        format="pdf",
                        category="layout",
                        size_bytes=sample_pdf_path.stat().st_size,
                    )
                ],
                target_format="pdf",
                action_name="split_pdf",
                options={"split_mode": split_mode},
                output_policy=OutputPolicy(output_dir=str(output_dir)),
            )

            result = task_mgr.execute_single(request)

            assert result.success, f"unexpected error for {split_mode}: {result.error}"
            assert [Path(artifact.staging_path).parent for artifact in result.artifacts] == [output_dir] * len(
                result.artifacts
            )
            assert [Path(artifact.staging_path).name for artifact in result.artifacts] == expected[
                "artifact_suggested_names"
            ]
            assert [Path(artifact.staging_path).suffix.lower() for artifact in result.artifacts] == [".pdf"] * len(
                result.artifacts
            )
            assert [artifact.media_type for artifact in result.artifacts] == ["application/pdf"] * len(result.artifacts)
            assert [artifact.metadata.get("split_mode") for artifact in result.artifacts] == [split_mode] * len(
                result.artifacts
            )
            if split_mode == "every_page":
                assert [artifact.metadata.get("page") for artifact in result.artifacts] == [1, 2, 3, 4]
            else:
                assert [artifact.metadata.get("pages") for artifact in result.artifacts] == expected[
                    "artifact_metadata_pages"
                ]
            assert [Path(artifact.staging_path).stat().st_size > 0 for artifact in result.artifacts] == [True] * len(
                result.artifacts
            )
            assert [_pdf_page_texts(artifact.staging_path) for artifact in result.artifacts] == expected[
                "artifact_page_groups"
            ]
            assert any(d.code == "PDF-SPLIT-OK" for d in result.diagnostics)
            assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
            assert all(str(workspace_root) not in str(artifact.staging_path) for artifact in result.artifacts)

    def test_plugin_dispatches_split_pdf_every_page_action_matches_old_system_fixture(self, tmp_path: Path) -> None:
        """Every-page split should keep one artifact per source page in source order."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        expected = fixture["projects"]["docwen-current"]["split_every_page"]
        sample_pdf_path = _create_pdf_ops_fixture_inputs(tmp_path)["sample.pdf"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "pdf",
                options={"split_mode": "every_page"},
                action_name="split_pdf",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == expected["metrics_page_count"]
            assert [artifact.suggested_name for artifact in result.artifacts] == expected["artifact_suggested_names"]
            assert [_pdf_page_texts(artifact.staging_path) for artifact in result.artifacts] == expected[
                "artifact_page_groups"
            ]
            assert result.metrics.extra["split_mode"] == expected["metrics_split_mode"]
            assert result.metrics.extra["page_count"] == expected["metrics_page_count"]

    def test_plugin_dispatches_split_pdf_odd_even_action_matches_old_system_fixture(self, tmp_path: Path) -> None:
        """Odd/even split should keep old-system odd-first, even-second grouping."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        expected = fixture["projects"]["docwen-current"]["split_odd_even"]
        sample_pdf_path = _create_pdf_ops_fixture_inputs(tmp_path)["sample.pdf"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "pdf",
                options={"split_mode": "odd_even"},
                action_name="split_pdf",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 2
            assert [artifact.suggested_name for artifact in result.artifacts] == expected["artifact_suggested_names"]
            assert [artifact.metadata.get("pages") for artifact in result.artifacts] == expected[
                "artifact_metadata_pages"
            ]
            assert [_pdf_page_texts(artifact.staging_path) for artifact in result.artifacts] == expected[
                "artifact_page_groups"
            ]
            assert result.metrics.extra["split_mode"] == expected["metrics_split_mode"]
            assert result.metrics.extra["odd_pages"] == expected["metrics_odd_pages"]
            assert result.metrics.extra["even_pages"] == expected["metrics_even_pages"]

    def test_plugin_rejects_single_page_pdf_split_every_page(self, tmp_path: Path) -> None:
        """Single-page PDFs have no useful split operation, matching old-system semantics."""
        from docwen_plugin_layout import LayoutPlugin

        fixture = _load_pdf_ops_old_system_fixture()
        expected = fixture["projects"]["docwen-current"]["split_single_page_every_page"]
        sample_pdf_path = _create_pdf_ops_fixture_inputs(tmp_path)["one.pdf"]

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "pdf",
                options={"split_mode": "every_page"},
                action_name="split_pdf",
            )
            result = LayoutPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == expected["error_type"]
            assert result.error.diagnostic_code == expected["diagnostic_code"]

    def test_can_handle_epub_is_false(self) -> None:
        """EPUB is no longer handled by the layout plugin."""
        from docwen_plugin_layout import LayoutPlugin

        plugin = LayoutPlugin()
        assert plugin.can_handle("epub", "md") is False
        assert plugin.can_handle("epub", "png") is False
        assert plugin.can_handle("epub", "pdf") is False

    def test_can_handle_ofd_xps_targets(self) -> None:
        """OFD/XPS can_handle must return True for all reachable target families."""
        from docwen_plugin_layout import LayoutPlugin

        plugin = LayoutPlugin()
        sources = ("ofd", "xps")
        targets = ("md", "png", "jpg", "tif", "docx", "doc", "odt", "rtf", "pdf")

        for src in sources:
            for tgt in targets:
                assert plugin.can_handle(src, tgt) is True, f"can_handle({src}, {tgt}) should be True"

    def test_retired_unsupported_sources_have_no_routes(self) -> None:
        from docwen_plugin_layout import LayoutPlugin

        plugin = LayoutPlugin()
        targets = ("md", "png", "jpg", "tif", "docx", "doc", "odt", "rtf", "pdf")

        for source in ("caj", "oxps", "layout"):
            for target in targets:
                assert plugin.can_handle(source, target) is False
