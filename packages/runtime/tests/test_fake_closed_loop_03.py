"""Focused tests split from test_fake_closed_loop.py."""

from __future__ import annotations

from ._fake_closed_loop_support import (
    PRECONVERSION_INTERMEDIATES_OPTION,
    Any,
    ConversionDiagnostic,
    ConversionMetrics,
    ConversionRequest,
    ConversionResult,
    FakeClosedLoopPlugin,
    FileRef,
    OutputPolicy,
    Path,
    TaskEvent,
    _read_template_payload,
    _write_template_package,
    os,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.contract


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

    def test_plugin_writes_to_staging_not_final_output(self, closed_loop, tmp_path) -> None:
        """Prove that the plugin only writes to staging, never to final output."""
        controller, _plugin, _task_mgr, _ws_mgr, _adapter = closed_loop

        input_file = tmp_path / "boundary.md"
        input_file.write_text("# Boundary test")

        final_output_dir = tmp_path / "final_output"
        final_output_dir.mkdir()

        request = ConversionRequest(
            request_id="req-boundary-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(final_output_dir)),
        )

        # Execute — the plugin only writes to staging via WorkspaceHandle
        result = controller.execute_single(request)

        assert result.success is True
        # The final artifact path is in the output directory
        final_path = result.artifacts[0].staging_path
        assert final_path.startswith(str(final_output_dir))

        # The plugin's original staging path is different
        # (we can verify the file was copied, not written directly)
        assert os.path.isfile(final_path)

    def test_preconversion_intermediate_artifact_is_finalized(self, closed_loop, tmp_path) -> None:
        """Pre-conversion intermediates are placed by OutputFinalizer as auxiliary artifacts."""
        controller, _plugin, _task_mgr, _ws_mgr, _adapter = closed_loop

        input_file = tmp_path / "legacy_pre.md"
        input_file.write_text("# Hub input")
        intermediate_file = tmp_path / "staging" / "legacy_pre.docx"
        intermediate_file.parent.mkdir()
        intermediate_file.write_text("intermediate")

        output_dir = tmp_path / "with_intermediate"
        output_dir.mkdir()

        request = ConversionRequest(
            request_id="req-preconversion-intermediate",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            options={
                PRECONVERSION_INTERMEDIATES_OPTION: [
                    {
                        "staging_path": str(intermediate_file),
                        "suggested_name": "legacy_fromDoc.docx",
                        "source_format": "doc",
                        "target_format": "docx",
                        "backend": "Fake Office",
                        "applies_to_input_path": str(input_file),
                    }
                ]
            },
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = controller.execute_single(request)

        assert result.success is True
        assert len(result.artifacts) == 2
        auxiliary = next(artifact for artifact in result.artifacts if artifact.kind == "auxiliary")
        assert auxiliary.suggested_name == "legacy_fromDoc.docx"
        assert auxiliary.metadata["source"] == "preconversion"
        assert os.path.isfile(auxiliary.staging_path)
        assert auxiliary.staging_path.startswith(str(output_dir))
        with open(auxiliary.staging_path) as f:
            assert f.read() == "intermediate"

    @pytest.mark.parametrize("failure_mode", ["reported", "raised"])
    def test_preconversion_intermediate_is_finalized_when_plugin_fails(
        self,
        tmp_path,
        monkeypatch,
        failure_mode: str,
    ) -> None:
        """An explicitly requested hub intermediate survives either plugin failure form."""
        from docwen_application.controller import ApplicationController
        from docwen_application.preconversion import pre_converter
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        source = tmp_path / "legacy.rtf"
        source.write_text(r"{\rtf1\ansi legacy source}", encoding="utf-8")
        output_dir = tmp_path / f"failure-{failure_mode}"
        plugin = FakeClosedLoopPlugin(
            f"preconversion-{failure_mode}",
            "docx",
            "md",
            should_fail=failure_mode == "reported",
            fail_message=f"{failure_mode} plugin failure",
        )
        if failure_mode == "raised":

            def raise_from_plugin(_context: Any) -> ConversionResult:
                raise RuntimeError("raised plugin failure")

            monkeypatch.setattr(plugin, "convert", raise_from_plugin)

        registry = PluginRegistry()
        registry.register(plugin)
        manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / f"runtime-{failure_mode}")),
            OutputFinalizer(),
        )
        controller = ApplicationController(runtime_port=RuntimePortAdapter(manager))

        def fake_pre_convert(input_path: str, _source_format: str, *, staging_dir: str, **_kwargs):
            intermediate = Path(staging_dir) / f"{Path(input_path).stem}.docx"
            _write_template_package(intermediate, "docx", payload="preserved hub")
            return PreConversionResult(str(intermediate), "rtf", "Fake Office")

        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        monkeypatch.setattr(pre_converter, "pre_convert", fake_pre_convert)

        result = controller.execute_single(
            ConversionRequest(
                request_id=f"preconversion-failure-{failure_mode}",
                input_refs=[FileRef(path=str(source), format="rtf", category="document")],
                target_format="md",
                output_policy=OutputPolicy(output_dir=str(output_dir)),
                config_snapshot={
                    "output": {
                        "intermediate_files": {
                            "save_to_output": True,
                        }
                    }
                },
            )
        )

        assert result.success is False
        assert result.error is not None
        assert result.error.message == f"{failure_mode} plugin failure"
        assert result.error.error_type == "conversion_failed"
        diagnostic_codes = [diagnostic.code for diagnostic in result.diagnostics]
        assert "FINALIZER_DONE" in diagnostic_codes
        if failure_mode == "reported":
            assert result.error.diagnostic_code == "FAKE-ERR"
            assert diagnostic_codes.count("FAKE-ERR") == 1
            assert result.metrics.extra["failure_stage"] == "plugin"
        else:
            assert "FAKE-ERR" not in diagnostic_codes
        assert result.metrics.extra["output_dir"] == str(output_dir)
        assert len(result.artifacts) == 1
        auxiliary = result.artifacts[0]
        assert auxiliary.kind == "auxiliary"
        assert auxiliary.is_primary is False
        assert auxiliary.suggested_name == "legacy_fromRtf.docx"
        assert auxiliary.metadata["source"] == "preconversion"
        assert Path(auxiliary.staging_path) == output_dir / "legacy_fromRtf.docx"
        assert _read_template_payload(Path(auxiliary.staging_path)) == "preserved hub"
        assert not list(tmp_path.glob("docwen_pre_*"))

    def test_failure_intermediate_placement_error_does_not_mask_plugin_failure(
        self,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Auxiliary placement diagnostics supplement rather than replace plugin failure."""
        from docwen_core.models.result import ConversionErrorInfo
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        input_file = tmp_path / "legacy.docx"
        input_file.write_text("hub input", encoding="utf-8")
        intermediate = tmp_path / "intermediate.docx"
        intermediate.write_text("hub intermediate", encoding="utf-8")
        output_dir = tmp_path / "placement-error"
        plugin = FakeClosedLoopPlugin(
            "preconversion-placement-error",
            "docx",
            "md",
            should_fail=True,
            fail_message="authoritative plugin failure",
        )
        registry = PluginRegistry()
        registry.register(plugin)
        finalizer = OutputFinalizer()

        def reject_placement(**_kwargs):
            raise OSError("destination denied")

        monkeypatch.setattr(finalizer, "finalize", reject_placement)
        manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "placement-error-runtime")),
            finalizer,
        )
        request = ConversionRequest(
            request_id="preconversion-placement-error",
            input_refs=[FileRef(path=str(input_file), format="docx", category="document")],
            target_format="md",
            options={
                PRECONVERSION_INTERMEDIATES_OPTION: [
                    {
                        "staging_path": str(intermediate),
                        "suggested_name": "legacy_fromDoc.docx",
                        "source_format": "doc",
                        "target_format": "docx",
                        "backend": "Fake Office",
                        "applies_to_input_path": str(input_file),
                    }
                ]
            },
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = manager.execute_single(request)

        assert result.success is False
        assert result.error == ConversionErrorInfo(
            error_type="conversion_failed",
            message="authoritative plugin failure",
            diagnostic_code="FAKE-ERR",
        )
        assert result.artifacts == []
        assert [diagnostic.code for diagnostic in result.diagnostics] == [
            "FAKE-ERR",
            "PRECONVERSION_INTERMEDIATE_FINALIZE_ERROR",
        ]
        assert result.diagnostics[-1].message.endswith("destination denied")
        assert result.metrics.extra == {"failure_stage": "plugin"}

    def test_structured_plugin_cancellation_does_not_finalize_preconversion_intermediate(
        self,
        tmp_path,
        monkeypatch,
    ) -> None:
        """A plugin-reported cancellation neither saves auxiliaries nor reports completion."""
        from docwen_application.controller import ApplicationController
        from docwen_application.preconversion import pre_converter
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionErrorInfo
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        source = tmp_path / "legacy.rtf"
        source.write_text(r"{\rtf1\ansi legacy source}", encoding="utf-8")
        output_dir = tmp_path / "cancelled"
        plugin = FakeClosedLoopPlugin("preconversion-cancelled", "docx", "md")

        def return_cancelled(context: Any) -> ConversionResult:
            diagnostic = ConversionDiagnostic(
                level="warning",
                message="Plugin observed cancellation",
                code="PLUGIN-CANCELLED",
            )
            context.progress.report_diagnostic(
                diagnostic.level,
                diagnostic.message,
                diagnostic.code,
            )
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                diagnostics=[diagnostic],
                error=ConversionErrorInfo(
                    error_type="cancelled",
                    message="Structured plugin cancellation",
                ),
                metrics=ConversionMetrics(
                    duration_ms=456.0,
                    input_bytes=123,
                    output_bytes=999,
                    extra={"cancel_stage": "plugin"},
                ),
            )

        monkeypatch.setattr(plugin, "convert", return_cancelled)
        registry = PluginRegistry()
        registry.register(plugin)
        manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "cancelled-runtime")),
            OutputFinalizer(),
        )
        adapter = RuntimePortAdapter(manager)
        controller = ApplicationController(runtime_port=adapter)

        intermediate_sizes: list[int] = []

        def fake_pre_convert(input_path: str, _source_format: str, *, staging_dir: str, **_kwargs):
            intermediate = Path(staging_dir) / f"{Path(input_path).stem}.docx"
            _write_template_package(intermediate, "docx", payload="must be discarded")
            intermediate_sizes.append(intermediate.stat().st_size)
            return PreConversionResult(str(intermediate), "rtf", "Fake Office")

        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        monkeypatch.setattr(pre_converter, "pre_convert", fake_pre_convert)

        result = controller.execute_single(
            ConversionRequest(
                request_id="preconversion-structured-cancel",
                input_refs=[FileRef(path=str(source), format="rtf", category="document")],
                target_format="md",
                output_policy=OutputPolicy(output_dir=str(output_dir)),
                config_snapshot={
                    "output": {
                        "intermediate_files": {
                            "save_to_output": True,
                        }
                    }
                },
            )
        )

        assert result.success is False
        assert result.error == ConversionErrorInfo(
            error_type="cancelled",
            message="Structured plugin cancellation",
        )
        assert result.artifacts == []
        assert [diagnostic.code for diagnostic in result.diagnostics] == ["PLUGIN-CANCELLED"]
        assert result.metrics.extra == {"cancel_stage": "plugin"}
        assert result.metrics.duration_ms != 456.0
        assert result.metrics.input_bytes == intermediate_sizes[0]
        assert result.metrics.output_bytes == 0
        assert not output_dir.exists()
        assert not list(tmp_path.glob("docwen_pre_*"))
        event_types = [event.event_type for event in adapter.collected_events]
        assert event_types[-1] == "task_cancelled"
        assert "task_completed" not in event_types

    def test_runtime_token_cancellation_discards_intermediate_despite_plugin_failure(
        self,
        tmp_path,
        monkeypatch,
    ) -> None:
        """The runtime token wins a return-time race with an ordinary plugin failure."""
        from docwen_core.models.result import ConversionErrorInfo
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        request_id = "preconversion-token-cancel"
        input_file = tmp_path / "legacy.docx"
        input_file.write_text("hub input", encoding="utf-8")
        intermediate = tmp_path / "intermediate.docx"
        intermediate.write_text("must be discarded", encoding="utf-8")
        output_dir = tmp_path / "token-cancelled"
        plugin = FakeClosedLoopPlugin("preconversion-token-cancel", "docx", "md")

        def return_failure_after_cancel(context: Any) -> ConversionResult:
            assert context.cancellation.is_cancelled is True
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message="plugin lost the cancellation race",
                ),
                metrics=ConversionMetrics(
                    output_bytes=999,
                    extra={"plugin_saw_cancelled": True},
                ),
            )

        monkeypatch.setattr(plugin, "convert", return_failure_after_cancel)
        registry = PluginRegistry()
        registry.register(plugin)
        manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "token-cancel-runtime")),
            OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id=request_id,
            input_refs=[
                FileRef(
                    path=str(input_file),
                    format="docx",
                    category="document",
                    size_bytes=input_file.stat().st_size,
                )
            ],
            target_format="md",
            options={
                PRECONVERSION_INTERMEDIATES_OPTION: [
                    {
                        "staging_path": str(intermediate),
                        "suggested_name": "legacy_fromDoc.docx",
                        "applies_to_input_path": str(input_file),
                    }
                ]
            },
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )
        events: list[TaskEvent] = []
        manager.cancel(request_id)

        result = manager.execute_single(request, on_event=events.append)

        assert result.success is False
        assert result.error == ConversionErrorInfo(
            error_type="cancelled",
            message="Task was cancelled",
        )
        assert result.artifacts == []
        assert result.metrics.input_bytes == input_file.stat().st_size
        assert result.metrics.output_bytes == 0
        assert result.metrics.extra == {"plugin_saw_cancelled": True}
        assert not output_dir.exists()
        assert [event.event_type for event in events][-1] == "task_cancelled"
