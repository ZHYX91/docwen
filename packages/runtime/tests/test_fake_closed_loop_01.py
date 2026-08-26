"""Focused tests split from test_fake_closed_loop.py."""

from __future__ import annotations

from ._fake_closed_loop_support import (
    ApplicationController,
    CapturingClosedLoopPlugin,
    ConversionRequest,
    FakeClosedLoopPlugin,
    FileRef,
    OutputPolicy,
    Path,
    StreamedWarningPlugin,
    _write_template_package,
    os,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.contract


def test_task_manager_docstring_has_no_stale_worker_phase_marker() -> None:
    """TaskManager is synchronous orchestration; caller-level pools live outside it."""
    import inspect

    from docwen_runtime.engine.task_manager import TaskManager

    source = inspect.getsource(TaskManager)
    assert "phase 4" not in source.lower()
    assert "workers/" not in source
    assert "caller-level thread/process pools" in source


class TestFullClosedLoop:
    """End-to-end: application command → workflow → adapter → runtime → result."""

    @pytest.fixture
    def closed_loop(self, tmp_path: str):
        """Build a complete closed loop with a fake plugin."""
        import tempfile

        plugin = FakeClosedLoopPlugin("fake_md2docx", "markdown", "docx")

        from docwen_application.controller import ApplicationController
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        ws_root = tempfile.mkdtemp()
        registry = PluginRegistry()
        registry.register(plugin)

        resolver = RouteResolver(registry)
        ws_mgr = WorkspaceManager(root_dir=ws_root)
        finalizer = OutputFinalizer()
        task_mgr = TaskManager(registry, resolver, ws_mgr, finalizer)
        adapter = RuntimePortAdapter(task_mgr)
        controller = ApplicationController(runtime_port=adapter)

        yield controller, plugin, task_mgr, ws_mgr, adapter

        # Cleanup
        ws_mgr.cleanup_all()
        import shutil

        shutil.rmtree(ws_root, ignore_errors=True)

    def test_single_file_success_flow(self, closed_loop, tmp_path) -> None:
        """Full closed loop: single file markdown→docx succeeds."""
        controller, _plugin, _task_mgr, _ws_mgr, _adapter = closed_loop

        # Create a real input file
        input_file = tmp_path / "test.md"
        input_file.write_text("# Hello\n\nWorld")

        output_dir = tmp_path / "output"
        output_dir.mkdir()

        request = ConversionRequest(
            request_id="req-success-001",
            input_refs=[
                FileRef(
                    path=str(input_file),
                    format="markdown",
                    category="markdown",
                    size_bytes=input_file.stat().st_size,
                )
            ],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = controller.execute_single(request)

        assert result.success is True
        assert len(result.artifacts) == 1
        # The artifact should be in the output directory
        assert result.artifacts[0].staging_path.startswith(str(output_dir))
        assert os.path.isfile(result.artifacts[0].staging_path)
        assert result.artifacts[0].suggested_name == "test.docx"
        assert result.metrics.extra["plugin_metric"] == "preserved"
        assert result.metrics.extra["output_dir"] == str(output_dir)

    def test_enabled_manifest_runs_through_application_runtime_and_finalizer(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_application.controller import ApplicationController
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.output.manifest import OutputManifestWriter
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        private_temp = tmp_path / "private-temp"
        private_temp.mkdir()
        monkeypatch.setattr(tempfile, "tempdir", str(private_temp))
        plugin = FakeClosedLoopPlugin("fake_manifest", "markdown", "docx")
        registry = PluginRegistry()
        registry.register(plugin)
        finalizer = OutputFinalizer()
        manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "workspace")),
            finalizer,
        )
        controller = ApplicationController(
            runtime_port=RuntimePortAdapter(
                manager,
                output_manifest_writer=OutputManifestWriter(finalizer),
            )
        )
        source = tmp_path / "private-source.md"
        source.write_text("# Secret", encoding="utf-8")
        output_dir = tmp_path / "output"
        request = ConversionRequest(
            request_id="manifest-closed-loop",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
            config_snapshot={
                "output": {
                    "manifest": {
                        "save_to_output": True,
                        "mask_input_path": True,
                    }
                }
            },
        )

        result = controller.execute_single(request)

        assert result.success is True
        assert [item.kind for item in result.artifacts] == ["primary", "manifest"]
        manifest_path = Path(result.artifacts[-1].staging_path)
        manifest_text = manifest_path.read_text(encoding="utf-8")
        assert manifest_path.parent == output_dir.resolve()
        assert "<redacted>/private-source.md" in manifest_text
        assert str(source.resolve()) not in manifest_text
        # Request-scoped manifest staging must be gone.  OutputFinalizer keeps
        # one process-lock sentinel per resolved output directory on purpose:
        # unlinking that path after release can split concurrent processes
        # across different inodes and defeat the cross-process lock.
        assert not any(path.name.startswith("docwen_manifest_") for path in private_temp.iterdir())
        lock_root = private_temp / "docwen-output-finalizer-locks"
        assert lock_root.is_dir()
        lock_entries = list(lock_root.iterdir())
        assert len(lock_entries) == 1
        assert lock_entries[0].suffix == ".lock"
        assert lock_entries[0].read_bytes() == b"\0"

    def test_read_only_policy_discards_staging_artifacts_without_finalization(self, closed_loop, tmp_path) -> None:
        controller, _plugin, _task_mgr, ws_mgr, _adapter = closed_loop
        input_file = tmp_path / "diagnostic.md"
        input_file.write_text("# Diagnostic", encoding="utf-8")
        output_dir = tmp_path / "must-not-exist"

        result = controller.execute_single(
            ConversionRequest(
                request_id="req-read-only-001",
                input_refs=[
                    FileRef(
                        path=str(input_file),
                        format="markdown",
                        category="markdown",
                        size_bytes=input_file.stat().st_size,
                    )
                ],
                target_format="docx",
                output_policy=OutputPolicy(output_dir=str(output_dir), write_artifacts=False),
            )
        )

        assert result.success is True
        assert result.artifacts == []
        assert result.metrics.output_bytes == 0
        assert result.metrics.extra["plugin_metric"] == "preserved"
        assert not output_dir.exists()
        assert ws_mgr.get("req-read-only-001") is None

    def test_template_canonical_id_is_resolved_to_resource_path_before_plugin_execution(
        self,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Runtime resolves the published template ID while plugins consume plain paths."""
        from docwen_application.controller import ApplicationController
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.templates import TemplateRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        templates_dir = tmp_path / "templates"
        templates_dir.mkdir()
        template_path = templates_dir / "Corporate Report.docx"
        _write_template_package(template_path, "docx")
        template_registry = TemplateRegistry(templates_dir)
        template_id = template_registry.list_templates()[0].id
        decoy_path = tmp_path / template_id
        decoy_path.write_bytes(b"not a template")
        monkeypatch.chdir(tmp_path)

        monkeypatch.setattr(
            TemplateRegistry,
            "default",
            staticmethod(lambda: template_registry),
        )

        plugin = CapturingClosedLoopPlugin("fake_md2docx", "markdown", "docx")
        registry = PluginRegistry()
        registry.register(plugin)

        resolver = RouteResolver(registry)
        ws_mgr = WorkspaceManager(root_dir=str(tmp_path / "workspace"))
        task_mgr = TaskManager(registry, resolver, ws_mgr, OutputFinalizer())
        controller = ApplicationController(runtime_port=RuntimePortAdapter(task_mgr))

        input_file = tmp_path / "note.md"
        input_file.write_text("# Hello\n\nWorld", encoding="utf-8")
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        request = ConversionRequest(
            request_id="req-template-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            action_name="",
            options={"template_name": template_id},
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = controller.execute_single(request)

        assert result.success
        assert plugin.seen_options
        assert plugin.seen_options[0]["template_name"] == str(template_path.resolve())
        assert decoy_path.read_bytes() == b"not a template"

    def test_xlsx_template_id_is_resolved_to_resource_path_before_plugin_execution(
        self,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Runtime resolves the published XLSX ID for Markdown spreadsheet routes."""
        from docwen_application.controller import ApplicationController
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.templates import TemplateRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        templates_dir = tmp_path / "templates"
        templates_dir.mkdir()
        docx_template_path = templates_dir / "Corporate Report.docx"
        xlsx_template_path = templates_dir / "Corporate Report.xlsx"
        _write_template_package(docx_template_path, "docx")
        _write_template_package(xlsx_template_path, "xlsx")

        template_registry = TemplateRegistry(templates_dir)
        xlsx_template_id = next(
            template.id for template in template_registry.list_templates() if template.target == "xlsx"
        )
        monkeypatch.setattr(
            TemplateRegistry,
            "default",
            staticmethod(lambda: template_registry),
        )

        plugin = CapturingClosedLoopPlugin("fake_md2xlsx", "markdown", "xlsx")
        registry = PluginRegistry()
        registry.register(plugin)

        resolver = RouteResolver(registry)
        ws_mgr = WorkspaceManager(root_dir=str(tmp_path / "workspace"))
        task_mgr = TaskManager(registry, resolver, ws_mgr, OutputFinalizer())
        controller = ApplicationController(runtime_port=RuntimePortAdapter(task_mgr))

        input_file = tmp_path / "note.md"
        input_file.write_text("# Hello\n\n| A |\n|---|\n| 1 |\n", encoding="utf-8")
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        request = ConversionRequest(
            request_id="req-xlsx-template-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="xlsx",
            action_name="",
            options={"template_name": xlsx_template_id},
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = controller.execute_single(request)

        assert result.success
        assert plugin.seen_options
        assert plugin.seen_options[0]["template_name"] == str(xlsx_template_path.resolve())

    def test_direct_template_path_is_rejected_before_plugin_execution(
        self,
        tmp_path,
    ) -> None:
        from docwen_application.controller import ApplicationController
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        template_path = tmp_path / "renamed-template.docx"
        _write_template_package(template_path, "docx")

        plugin = CapturingClosedLoopPlugin("fake_md2docx", "markdown", "docx")
        registry = PluginRegistry()
        registry.register(plugin)
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "workspace")),
            OutputFinalizer(),
        )
        controller = ApplicationController(runtime_port=RuntimePortAdapter(task_mgr))
        input_file = tmp_path / "note.md"
        input_file.write_text("# Hello", encoding="utf-8")

        result = controller.execute_single(
            ConversionRequest(
                request_id="req-template-path-rejected",
                input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
                target_format="docx",
                options={"template_name": str(template_path)},
                output_policy=OutputPolicy(output_dir=str(tmp_path / "output")),
            )
        )

        assert result.success is False
        assert plugin.seen_options == []
        assert result.error is not None
        assert result.error.error_type == "invalid_input"
        assert result.error.diagnostic_code == "TEMPLATE_ID_INVALID"

    def test_pdf_to_docx_uses_pdf2docx_fallback_through_runtime(
        self,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Runtime route execution reaches the PDF→DOCX pdf2docx fallback."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        input_file = tmp_path / "sample.pdf"
        input_file.write_bytes(b"%PDF-1.4\n% fake pdf is enough for mocked conversion\n")
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        observed_priority: list[str] = []

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, output_path, cancel, failure_subject
            assert source_format == "pdf"
            observed_priority.extend(backend_priority)
            assert set(com_candidates) == {"msoffice_word"}
            assert libreoffice_format == "docx"
            return BridgeResult(False, message="Office and LibreOffice unavailable")

        def fake_pdf2docx(input_path, output_path, *, cancellation=None):
            assert cancellation is not None
            output_path.write_bytes(b"docx-by-pdf2docx-runtime")
            return BridgeResult(True, output_path=str(output_path), backend="pdf2docx")

        monkeypatch.setattr(converter, "convert_with_backend_priority", fake_convert_with_backend_priority)
        monkeypatch.setattr(converter, "_convert_pdf_with_pdf2docx", fake_pdf2docx)

        registry = PluginRegistry()
        registry.register(LayoutPlugin())
        resolver = RouteResolver(registry)
        ws_mgr = WorkspaceManager(root_dir=str(tmp_path / "workspace"))
        task_mgr = TaskManager(registry, resolver, ws_mgr, OutputFinalizer())
        controller = ApplicationController(runtime_port=RuntimePortAdapter(task_mgr))

        request = ConversionRequest(
            request_id="req-pdf-docx-fallback",
            input_refs=[FileRef(path=str(input_file), format="pdf", category="layout")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = controller.execute_single(request)

        assert result.success is True, f"unexpected error: {result.error}"
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        assert artifact.suggested_name == "sample.docx"
        assert artifact.staging_path.startswith(str(output_dir))
        assert os.path.isfile(artifact.staging_path)
        with open(artifact.staging_path, "rb") as fh:
            assert fh.read() == b"docx-by-pdf2docx-runtime"
        assert observed_priority == ["msoffice_word", "libreoffice"]

    def test_single_file_failure(self, closed_loop, tmp_path: Path) -> None:
        """Full closed loop: plugin reports failure."""
        _controller, _, _, _, _adapter = closed_loop

        # Replace with a failing plugin
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        failing_plugin = FakeClosedLoopPlugin("failing", "markdown", "docx", should_fail=True)
        ws_root2 = tmp_path / "failure-workspace"
        ws_root2.mkdir()
        reg2 = PluginRegistry()
        reg2.register(failing_plugin)
        resolver2 = RouteResolver(reg2)
        ws_mgr2 = WorkspaceManager(root_dir=str(ws_root2))
        tm2 = TaskManager(reg2, resolver2, ws_mgr2, OutputFinalizer())
        adapter2 = RuntimePortAdapter(tm2)
        controller2 = ApplicationController(runtime_port=adapter2)

        input_file = tmp_path / "fail.md"
        input_file.write_text("# Will fail")

        request = ConversionRequest(
            request_id="req-fail-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        result = controller2.execute_single(request)
        assert result.success is False
        assert result.metrics.extra["failure_stage"] == "plugin"
        assert [diagnostic.code for diagnostic in result.diagnostics] == ["FAKE-ERR"]

    def test_streamed_diagnostic_is_merged_and_deduplicated(self, tmp_path: Path) -> None:
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        plugin = StreamedWarningPlugin("streamed_warning", "markdown", "docx")
        registry = PluginRegistry()
        registry.register(plugin)
        task_manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "workspace")),
            OutputFinalizer(),
        )
        input_file = tmp_path / "sample.md"
        input_file.write_text("# Sample")
        request = ConversionRequest(
            request_id="req-streamed-warning",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        result = task_manager.execute_single(request)

        assert result.success is True
        warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.code == "OCR-BEST-EFFORT"]
        assert len(warnings) == 1
        assert warnings[0].location == "sample.png"
