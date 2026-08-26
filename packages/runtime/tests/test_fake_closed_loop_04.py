"""Focused tests split from test_fake_closed_loop.py."""

from __future__ import annotations

from ._fake_closed_loop_support import (
    PRECONVERSION_INTERMEDIATES_OPTION,
    ApplicationController,
    ConversionRequest,
    ConversionResult,
    FakeClosedLoopPlugin,
    FileRef,
    OutputPolicy,
    Path,
    PreconversionIdentityPlugin,
    TaskEvent,
    _write_template_package,
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

    def test_failure_intermediate_finalizes_once_and_listener_cannot_mask_plugin_error(
        self,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Failure placement is single-shot and terminal listeners cannot replace its error."""
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
        output_dir = tmp_path / "single-shot"
        plugin = FakeClosedLoopPlugin(
            "preconversion-single-shot",
            "docx",
            "md",
            should_fail=True,
            fail_message="authoritative plugin failure",
        )
        registry = PluginRegistry()
        registry.register(plugin)
        finalizer = OutputFinalizer()
        real_finalize = finalizer.finalize
        finalize_calls: list[str] = []

        def count_finalize(**kwargs):
            finalize_calls.append(kwargs["task_id"])
            return real_finalize(**kwargs)

        monkeypatch.setattr(finalizer, "finalize", count_finalize)
        manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "single-shot-runtime")),
            finalizer,
        )
        request = ConversionRequest(
            request_id="preconversion-single-shot",
            input_refs=[FileRef(path=str(input_file), format="docx", category="document")],
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

        def rejecting_listener(event: TaskEvent) -> None:
            events.append(event)
            if event.event_type in {"task_completed", "task_failed"}:
                raise RuntimeError("presenter callback failed")

        result = manager.execute_single(request, on_event=rejecting_listener)

        assert result.success is False
        assert result.error == ConversionErrorInfo(
            error_type="conversion_failed",
            message="authoritative plugin failure",
            diagnostic_code="FAKE-ERR",
        )
        assert finalize_calls == [request.request_id]
        assert [path.name for path in output_dir.iterdir()] == ["legacy_fromDoc.docx"]
        assert [artifact.suggested_name for artifact in result.artifacts] == ["legacy_fromDoc.docx"]
        assert "TASK_EVENT_LISTENER_ERROR" in [diagnostic.code for diagnostic in result.diagnostics]
        event_types = [event.event_type for event in events]
        assert len([event for event in event_types if event in {"task_completed", "task_failed"}]) == 1

    def test_preconversion_same_stem_batch_finalizes_to_each_original_parent(
        self,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Application staging is isolated, source-anchored and gone after finalization."""
        from docwen_application.controller import ApplicationController
        from docwen_application.preconversion import pre_converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_runtime.adapters import RuntimePortAdapter
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        first_dir = tmp_path / "first"
        second_dir = tmp_path / "second"
        first_dir.mkdir()
        second_dir.mkdir()
        first = first_dir / "report.rtf"
        second = second_dir / "report.rtf"
        first_content = r"{\rtf1\ansi FIRST}"
        second_content = r"{\rtf1\ansi SECOND}"
        first.write_text(first_content, encoding="utf-8")
        second.write_text(second_content, encoding="utf-8")

        plugin = PreconversionIdentityPlugin()
        registry = PluginRegistry()
        registry.register(plugin)
        manager = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "runtime-workspaces")),
            OutputFinalizer(),
        )
        controller = ApplicationController(runtime_port=RuntimePortAdapter(manager))
        protected_inputs: list[tuple[Path, str]] = []

        def fake_bridge(input_path: str, output_path: str, **_kwargs):
            protected = Path(input_path)
            protected_inputs.append((protected, protected.read_text(encoding="utf-8")))
            _write_template_package(
                Path(output_path),
                "docx",
                payload=protected.read_text(encoding="utf-8"),
            )
            return BridgeResult(True, output_path=output_path, backend="Fake Office")

        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        monkeypatch.setattr(pre_converter, "convert_with_backend_priority", fake_bridge)

        results = controller.execute_batch(
            ConversionRequest(
                request_id="same-stem-closed-loop",
                input_refs=[
                    FileRef(path=str(first), format="rtf", category="document"),
                    FileRef(path=str(second), format="rtf", category="document"),
                ],
                target_format="md",
                output_policy=OutputPolicy(),
            )
        )

        assert [result.success for result in results] == [True, True]
        assert [path.name for path, _content in protected_inputs] == ["input.rtf", "input.rtf"]
        assert [content for _path, content in protected_inputs] == [first_content, second_content]
        assert protected_inputs[0][0].parent != protected_inputs[1][0].parent
        assert [content for _path, content in plugin.seen] == [first_content, second_content]
        assert plugin.seen[0][0] != plugin.seen[1][0]
        final_paths = [Path(result.artifacts[0].staging_path) for result in results]
        assert [path.parent.parent for path in final_paths] == [first_dir, second_dir]
        assert all(path.parent.name == path.stem for path in final_paths)
        assert all(path.stem.startswith("report_") and path.stem.endswith("_fromDocx") for path in final_paths)
        assert all((path.parent / "docwen-node.json").is_file() for path in final_paths)
        assert [path.read_text(encoding="utf-8") for path in final_paths] == [first_content, second_content]
        assert [first.read_text(encoding="utf-8"), second.read_text(encoding="utf-8")] == [
            first_content,
            second_content,
        ]
        assert all(not path.exists() for path, _content in protected_inputs)
        assert not list(tmp_path.glob("docwen_pre_*"))

    def test_task_events_are_collected_in_order(self, closed_loop, tmp_path) -> None:
        """Events collected during conversion must be in order."""
        _controller, _plugin, task_mgr, _ws_mgr, _adapter = closed_loop

        events: list[TaskEvent] = []

        input_file = tmp_path / "order.md"
        input_file.write_text("# Order")

        request = ConversionRequest(
            request_id="req-order-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        task_mgr.execute_single(request, on_event=lambda e: events.append(e))

        # Verify event sequence
        event_sequence = [(e.sequence, e.event_type) for e in events]
        assert event_sequence[0][1] == "task_started"
        assert event_sequence[-1][1] == "task_completed"

        # Progress events in between
        progress_events = [e for e in events if e.event_type == "task_progress"]
        assert len(progress_events) >= 1

    def test_result_contains_metrics(self, closed_loop, tmp_path) -> None:
        """ConversionResult must include timing and size metrics."""
        controller, _, _, _, _adapter = closed_loop

        input_file = tmp_path / "metrics.md"
        input_file.write_text("# Metrics test")

        request = ConversionRequest(
            request_id="req-metrics-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        result = controller.execute_single(request)
        assert result.metrics.duration_ms > 0
        assert result.metrics.output_bytes > 0

    def test_progress_events_reach_adapter(self, closed_loop, tmp_path) -> None:
        """Plugin-reported progress events must flow through to the adapter layer.

        Verifies the full event pipeline:
          Plugin → _RuntimeProgressSink → TaskManager → RuntimePortAdapter
        """
        _controller, _plugin, task_mgr, _ws_mgr, _adapter = closed_loop

        # Wire up the adapter with event collection
        from docwen_runtime.adapters import RuntimePortAdapter

        adapter = RuntimePortAdapter(task_mgr)
        adapter_controller = ApplicationController(runtime_port=adapter)

        input_file = tmp_path / "progress_e2e.md"
        input_file.write_text("# Progress E2E")

        request = ConversionRequest(
            request_id="req-progress-e2e-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        result = adapter_controller.execute_single(request)
        assert result.success is True

        # Events must be collected by the adapter
        events = adapter.collected_events
        assert len(events) > 0, "No events collected — progress pipeline is broken"

        # Verify event types are TaskEvent instances (not raw dicts)
        from docwen_core.models.task import TaskEvent

        for evt in events:
            assert isinstance(evt, TaskEvent), f"Expected TaskEvent, got {type(evt).__name__}: {evt!r}"

        # Verify the key event types are present
        event_types = {e.event_type for e in events}
        assert "task_started" in event_types
        assert "task_completed" in event_types

        # Plugin-reported progress events must be present
        progress_events = [e for e in events if e.event_type == "task_progress"]
        assert len(progress_events) >= 1, (
            "Plugin progress events not reaching adapter — pipeline H-2/H-3 may still be broken"
        )

    def test_progress_events_are_typed_not_dicts(self, closed_loop, tmp_path) -> None:
        """All events flowing through the pipeline must be TaskEvent instances, not raw dicts.

        This is the regression test for H-1.
        """
        _controller, _plugin, task_mgr, _ws_mgr, _adapter = closed_loop

        events: list[TaskEvent] = []

        input_file = tmp_path / "typed_events.md"
        input_file.write_text("# Typed events")

        request = ConversionRequest(
            request_id="req-typed-events-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        result = task_mgr.execute_single(request, on_event=lambda e: events.append(e))
        assert result.success is True

        from docwen_core.models.task import TaskEvent

        for evt in events:
            assert isinstance(evt, TaskEvent), (
                f"H-1 regression: event is {type(evt).__name__}, not TaskEvent. Raw dict event: {evt!r}"
            )
            # Verify it has the required fields
            assert evt.task_id
            assert isinstance(evt.sequence, int)
            assert evt.event_type
            assert evt.timestamp

        # Plugin progress events specifically must be TaskEvent
        progress_events = [e for e in events if e.event_type == "task_progress"]
        assert len(progress_events) >= 1
        for pe in progress_events:
            assert "percent" in pe.payload, f"Progress event missing percent: {pe.payload}"


class TestBatchModelStates:
    """Model-level tests for batch success, failure, skip, cancel states."""

    def test_batch_all_success(self) -> None:
        """All items in a batch succeed."""
        results = [
            ConversionResult(task_id="t1", success=True),
            ConversionResult(task_id="t2", success=True),
            ConversionResult(task_id="t3", success=True),
        ]
        assert all(r.success for r in results)
        assert sum(1 for r in results if r.success) == 3

    def test_batch_all_failed(self) -> None:
        """All items in a batch fail."""
        from docwen_core.models.result import ConversionErrorInfo

        results = [
            ConversionResult(
                task_id="t1", success=False, error=ConversionErrorInfo(error_type="conversion_failed", message="err1")
            ),
            ConversionResult(
                task_id="t2", success=False, error=ConversionErrorInfo(error_type="conversion_failed", message="err2")
            ),
        ]
        assert sum(1 for r in results if r.success) == 0
        assert all(r.error is not None for r in results)

    def test_batch_mixed_success_failure(self) -> None:
        """Some succeed, some fail."""
        from docwen_core.models.result import ConversionErrorInfo

        results = [
            ConversionResult(task_id="t1", success=True),
            ConversionResult(
                task_id="t2", success=False, error=ConversionErrorInfo(error_type="conversion_failed", message="err")
            ),
            ConversionResult(task_id="t3", success=True),
        ]
        assert sum(1 for r in results if r.success) == 2
        assert sum(1 for r in results if not r.success) == 1

    def test_batch_skipped_items(self) -> None:
        """Items can be explicitly skipped."""
        from docwen_core.models.result import ConversionErrorInfo

        results = [
            ConversionResult(task_id="t1", success=True),
            ConversionResult(
                task_id="t2",
                success=False,
                error=ConversionErrorInfo(error_type="skipped", message="Skipped due to earlier failure"),
            ),
            ConversionResult(
                task_id="t3",
                success=False,
                error=ConversionErrorInfo(error_type="skipped", message="Skipped due to earlier failure"),
            ),
        ]
        skipped = [r for r in results if r.error and r.error.error_type == "skipped"]
        assert len(skipped) == 2

    def test_batch_cancelled_items(self) -> None:
        """Items can be marked as cancelled."""
        from docwen_core.models.result import ConversionErrorInfo

        results = [
            ConversionResult(task_id="t1", success=True),
            ConversionResult(
                task_id="t2", success=False, error=ConversionErrorInfo(error_type="cancelled", message="User cancelled")
            ),
        ]
        cancelled = [r for r in results if r.error and r.error.error_type == "cancelled"]
        assert len(cancelled) == 1

    def test_batch_summary_counts(self) -> None:
        """BatchWorkflow summary computes correct counts."""
        from docwen_core.models.result import ConversionErrorInfo

        results = [
            ConversionResult(task_id="t1", success=True),
            ConversionResult(task_id="t2", success=True),
            ConversionResult(
                task_id="t3", success=False, error=ConversionErrorInfo(error_type="conversion_failed", message="err")
            ),
            ConversionResult(
                task_id="t4", success=False, error=ConversionErrorInfo(error_type="skipped", message="skipped")
            ),
            ConversionResult(
                task_id="t5", success=False, error=ConversionErrorInfo(error_type="cancelled", message="cancelled")
            ),
        ]

        from docwen_application.workflows.batch import BatchWorkflow

        # BatchWorkflow needs a runtime port; we only test summary logic
        summary = BatchWorkflow.summary(None, results)  # type: ignore[arg-type]
        assert summary["total"] == 5
        assert summary["success"] == 2
        assert summary["failed"] == 1
        assert summary["skipped"] == 1
        assert summary["cancelled"] == 1
