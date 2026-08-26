"""Focused tests split from test_xlsx_to_md_golden.py."""

from __future__ import annotations

from ._xlsx_to_md_golden_support import (
    Any,
    Generator,
    Path,
    _deliverable_artifacts,
    _document_node_root,
    _load_xlsx_to_md_old_system_fixture,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.golden


@pytest.mark.integration
class TestSpreadsheetToMdPipeline:
    """Test spreadsheet→MD through the full runtime pipeline."""

    @pytest.fixture
    def pipeline(self) -> Generator[dict[str, Any], None, None]:
        """Build a full runtime pipeline for spreadsheet→md conversion."""
        from docwen_plugin_spreadsheet.plugin import SpreadsheetPlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        plugin = SpreadsheetPlugin()
        plugin_registry = PluginRegistry()
        plugin_registry.register(plugin)

        route_resolver = RouteResolver(plugin_registry)

        ws_root = tempfile.mkdtemp(prefix="docwen_test_ws_")
        ws_manager = WorkspaceManager(ws_root)
        finalizer = OutputFinalizer()
        task_manager = TaskManager(
            plugin_registry=plugin_registry,
            route_resolver=route_resolver,
            workspace_manager=ws_manager,
            output_finalizer=finalizer,
        )

        yield {
            "task_manager": task_manager,
            "workspace_manager": ws_manager,
            "ws_root": ws_root,
        }

        # Cleanup
        import shutil

        shutil.rmtree(ws_root, ignore_errors=True)

    @pytest.mark.parametrize(
        "case_name",
        [
            "utf8_simple.csv",
            "gbk_simple.csv",
            "semicolon_simple.csv",
            "two_blocks.tsv",
        ],
    )
    def test_pipeline_delimited_output_directory_batch_matches_old_system_projection(
        self, pipeline: dict[str, Any], tmp_path: Path, case_name: str
    ) -> None:
        """Focused CSV/TSV batch finalizes to output_dir with old-system semantics."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        fixture = _load_xlsx_to_md_old_system_fixture()
        scope = fixture["shared_behavior_evidence"]["delimited_output_directory_batch_scope"]
        case = next(item for item in scope["probe_cases"] if item["name"] == case_name)
        input_path = tmp_path / case["name"]
        input_path.write_text(case["text"], encoding=case["encoding"])
        output_dir = tmp_path / f"{input_path.stem}_out"
        output_dir.mkdir()

        request = ConversionRequest(
            request_id=f"pipe-delimited-output-{input_path.stem}",
            input_refs=[
                FileRef(
                    path=str(input_path),
                    format=case["actual_format"],
                    category="spreadsheet",
                    size_bytes=input_path.stat().st_size,
                )
            ],
            target_format="md",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = pipeline["task_manager"].execute_single(request)

        assert result.success is True
        assert [diagnostic.code for diagnostic in result.diagnostics] == scope["current_projection"]["diagnostic_codes"]
        assert len(_deliverable_artifacts(result)) == 1
        artifact = _deliverable_artifacts(result)[0]
        assert artifact.media_type == scope["current_projection"]["artifact_media_type"]
        assert artifact.metadata["source_suggested_name"] == case["expected_current_suggested_name"]
        artifact_path = Path(artifact.staging_path)
        node_root = _document_node_root(artifact_path, output_dir)
        assert artifact_path.name == f"{node_root.name}.md"
        assert artifact_path.is_file()

        content = artifact_path.read_text(encoding="utf-8")
        assert str(Path(pipeline["ws_root"])) not in content
        assert str(output_dir) not in content

        if case["actual_format"] == "csv":
            for key, value in scope["current_projection"]["csv_artifact_metadata"].items():
                assert artifact.metadata[key] == value
            assert (
                sum(1 for line in content.splitlines() if line.startswith("|"))
                == scope["current_projection"]["csv_table_line_count"]
            )
            assert (
                sum(1 for line in content.splitlines() if line.startswith("|:"))
                == scope["current_projection"]["csv_separator_line_count"]
            )
            for token in scope["required_csv_tokens"]:
                assert token in content
        else:
            for key, value in scope["current_projection"]["tsv_artifact_metadata"].items():
                assert artifact.metadata[key] == value
            assert (
                sum(1 for line in content.splitlines() if line.startswith("|"))
                == scope["current_projection"]["tsv_table_line_count"]
            )
            assert (
                sum(1 for line in content.splitlines() if line.startswith("|:"))
                == scope["current_projection"]["tsv_separator_line_count"]
            )
            for token in scope["required_tsv_tokens"]:
                assert token in content


@pytest.mark.contract
def test_plugin_does_not_depend_on_runtime() -> None:
    """Plugin module must not import runtime at the source level."""
    import sys

    forbidden = {"docwen_runtime", "docwen_application"}
    for mod_name in list(sys.modules):
        if not mod_name.startswith("docwen_plugin_spreadsheet"):
            continue
        if any(mod_name.startswith(f) for f in forbidden):
            continue  # test files may import runtime
        mod = sys.modules[mod_name]
        for attr_name in dir(mod):
            attr = getattr(mod, attr_name, None)
            if attr is None:
                continue
            attr_mod = getattr(attr, "__module__", "")
            for fb in forbidden:
                if attr_mod.startswith(fb):
                    raise AssertionError(f"Plugin module {mod_name}.{attr_name} imports {attr_mod}")
