"""Focused tests split from test_fake_closed_loop.py."""

from __future__ import annotations

from ._fake_closed_loop_support import (
    _THREAD_COORDINATION_TIMEOUT_SECONDS,
    PRECONVERSION_INTERMEDIATES_OPTION,
    Any,
    ConversionRequest,
    ConversionResult,
    FakeClosedLoopPlugin,
    FileRef,
    OutputPolicy,
    Path,
    StreamedWarningThenRaisePlugin,
    TaskEvent,
    pytest,
    threading,
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

    @pytest.mark.parametrize(
        ("failure", "expected_error_type"),
        [
            pytest.param("cancel", "cancelled", id="cancellation"),
            pytest.param("runtime", "conversion_failed", id="unexpected-exception"),
        ],
    )
    def test_streamed_diagnostic_survives_plugin_exception(
        self,
        tmp_path: Path,
        failure: str,
        expected_error_type: str,
    ) -> None:
        from docwen_core.errors import CancellationRequested
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        raised = CancellationRequested("cancelled after warning") if failure == "cancel" else RuntimeError("boom")
        plugin = StreamedWarningThenRaisePlugin(raised)
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
            request_id=f"req-streamed-{failure}",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        result = task_manager.execute_single(request)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == expected_error_type
        assert [(item.code, item.location) for item in result.diagnostics] == [("OCR-BEST-EFFORT", "sample.png")]

    def test_event_ordering(self, closed_loop, tmp_path) -> None:
        """Task events must follow the correct order: started → progress → completed."""
        _controller, _plugin, task_mgr, _ws_mgr, _adapter = closed_loop

        events: list[TaskEvent] = []

        def collect(event: TaskEvent) -> None:
            events.append(event)

        input_file = tmp_path / "event_test.md"
        input_file.write_text("# Test")

        request = ConversionRequest(
            request_id="req-events-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        task_mgr.execute_single(request, on_event=collect)

        event_types = [e.event_type for e in events]
        assert "task_started" in event_types
        assert "task_progress" in event_types
        assert "task_completed" in event_types

        # Started must be first
        assert event_types[0] == "task_started"

        # Sequence numbers must be monotonically increasing
        sequences = [e.sequence for e in events]
        assert sequences == sorted(sequences)

    def test_batch_success(self, closed_loop, tmp_path) -> None:
        """Batch: multiple files, all succeed."""
        controller, _plugin, _task_mgr, _ws_mgr, _adapter = closed_loop

        files = []
        for i in range(3):
            f = tmp_path / f"file_{i}.md"
            f.write_text(f"# File {i}")
            files.append(f)

        output_dir = tmp_path / "batch_out"
        output_dir.mkdir()

        request = ConversionRequest(
            request_id="req-batch-001",
            input_refs=[FileRef(path=str(f), format="markdown", category="markdown") for f in files],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        results = controller.execute_batch(request)

        assert len(results) == 3
        for r in results:
            assert r.success is True
            assert len(r.artifacts) == 1

    def test_batch_partial_failure(self, closed_loop, tmp_path) -> None:
        """Batch: some files fail, others succeed (continue_on_error)."""
        # We test at task_manager level for partial failure
        _controller, _, task_mgr, _ws_mgr, _adapter = closed_loop

        # Build a custom scenario with mixed inputs
        output_dir = tmp_path / "batch_partial"
        output_dir.mkdir()

        # Use task_manager directly for more control
        # 3 files, all go through the same (working) plugin → all succeed
        files = []
        for i in range(3):
            f = tmp_path / f"partial_{i}.md"
            f.write_text(f"# File {i}")
            files.append(f)

        request = ConversionRequest(
            request_id="req-batch-partial-001",
            input_refs=[FileRef(path=str(f), format="markdown", category="markdown") for f in files],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        results = task_mgr.execute_batch(request, continue_on_error=True)
        assert len(results) == 3
        success_count = sum(1 for r in results if r.success)
        assert success_count == 3

    def test_batch_skip_on_stop(self, closed_loop, tmp_path: Path) -> None:
        """Batch: when continue_on_error=False, first failure stops the batch."""
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        # Create a plugin that fails on specific files
        ws_root = tmp_path / "batch-stop-workspace"
        ws_root.mkdir()
        failing_plugin = FakeClosedLoopPlugin("flaky", "markdown", "docx", should_fail=True)
        reg = PluginRegistry()
        reg.register(failing_plugin)
        resolver = RouteResolver(reg)
        ws_mgr = WorkspaceManager(root_dir=str(ws_root))
        tm = TaskManager(reg, resolver, ws_mgr, OutputFinalizer())

        files = []
        for i in range(3):
            f = tmp_path / f"stop_{i}.md"
            f.write_text(f"# File {i}")
            files.append(f)

        request = ConversionRequest(
            request_id="req-batch-stop-001",
            input_refs=[FileRef(path=str(f), format="markdown", category="markdown") for f in files],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "stop_out")),
        )

        results = tm.execute_batch(request, continue_on_error=False)

        # First file fails, remaining are skipped
        assert len(results) == 3
        assert results[0].success is False
        assert results[1].success is False
        assert results[1].error is not None
        assert results[1].error.error_type == "skipped", (
            f"Expected error_type='skipped', got {results[1].error.error_type!r}"
        )

    def test_cancellation_during_conversion(self, closed_loop, tmp_path) -> None:
        """Task cancellation before execution produces a cancelled result."""
        _controller, _plugin, task_mgr, _ws_mgr, _adapter = closed_loop

        input_file = tmp_path / "cancel.md"
        input_file.write_text("# Test")

        request = ConversionRequest(
            request_id="req-cancel-001",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        # Cancel before execution
        task_mgr.cancel("req-cancel-001")

        result = task_mgr.execute_single(request)
        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "cancelled"

    def test_cancel_after_finalizer_commit_keeps_completed_terminal(
        self,
        closed_loop,
        tmp_path,
        monkeypatch,
    ) -> None:
        """The finalizer commit point linearizes a later cancellation request."""
        _controller, _plugin, task_mgr, _ws_mgr, _adapter = closed_loop
        from docwen_runtime.output.finalizer import OutputFinalizer

        input_file = tmp_path / "late-finalizer-cancel.md"
        input_file.write_text("# committed", encoding="utf-8")
        output_dir = tmp_path / "late-finalizer-cancel-out"
        request = ConversionRequest(
            request_id="late-finalizer-cancel",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )
        events: list[TaskEvent] = []
        real_commit = OutputFinalizer._commit_prepared

        def commit_then_cancel(*args: Any, **kwargs: Any) -> Any:
            committed = real_commit(*args, **kwargs)
            task_mgr.cancel(request.request_id)
            return committed

        monkeypatch.setattr(OutputFinalizer, "_commit_prepared", staticmethod(commit_then_cancel))

        result = task_mgr.execute_single(request, on_event=events.append)

        assert result.success is True
        assert Path(result.artifacts[0].staging_path).is_file()
        assert [event.event_type for event in events][-1] == "task_completed"
        assert "task_cancelled" not in [event.event_type for event in events]

    def test_cancel_at_finalizer_precommit_emits_cancelled_without_output(
        self,
        closed_loop,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Cancellation wins at the final check before publication begins."""
        _controller, _plugin, task_mgr, _ws_mgr, _adapter = closed_loop
        from docwen_runtime.output.finalizer import OutputFinalizer

        input_file = tmp_path / "precommit-cancel.md"
        input_file.write_text("# prepared", encoding="utf-8")
        output_dir = tmp_path / "precommit-cancel-out"
        request = ConversionRequest(
            request_id="precommit-cancel",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )
        events: list[TaskEvent] = []
        real_prepare = OutputFinalizer._prepare_artifact

        def prepare_then_cancel(*args: Any, **kwargs: Any) -> Any:
            prepared = real_prepare(*args, **kwargs)
            task_mgr.cancel(request.request_id)
            return prepared

        monkeypatch.setattr(OutputFinalizer, "_prepare_artifact", staticmethod(prepare_then_cancel))

        result = task_mgr.execute_single(request, on_event=events.append)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "cancelled"
        assert [event.event_type for event in events][-1] == "task_cancelled"
        assert "task_completed" not in [event.event_type for event in events]
        assert list(output_dir.iterdir()) == []

    def test_application_reservation_closes_runtime_return_cancel_gap(
        self,
        closed_loop,
        tmp_path,
        monkeypatch,
    ) -> None:
        """A late click cannot recreate pending state after Runtime returns."""
        controller, _plugin, task_mgr, _ws_mgr, _adapter = closed_loop
        input_file = tmp_path / "return-gap.md"
        input_file.write_text("# Return gap", encoding="utf-8")
        request = ConversionRequest(
            request_id="return-gap",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "return-gap-out")),
        )
        runtime_returned = threading.Event()
        release_finish = threading.Event()
        real_finish = controller._finish_runtime_task
        results: list[ConversionResult] = []

        def paused_finish(scope: Any, task_id: str) -> None:
            runtime_returned.set()
            assert release_finish.wait(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            real_finish(scope, task_id)

        monkeypatch.setattr(controller, "_finish_runtime_task", paused_finish)
        worker = threading.Thread(target=lambda: results.append(controller.execute_single(request)))
        worker.start()
        try:
            assert runtime_returned.wait(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            assert task_mgr.active_tasks == []
            controller.cancel(request.request_id)
            assert task_mgr._pending_cancellations == set()
            release_finish.set()
            worker.join(_THREAD_COORDINATION_TIMEOUT_SECONDS)
        finally:
            release_finish.set()
            worker.join(_THREAD_COORDINATION_TIMEOUT_SECONDS)

        assert not worker.is_alive()
        assert len(results) == 1
        assert results[0].success is True
        assert task_mgr._tokens == {}
        assert task_mgr._reserved_cancellations == set()

    def test_reserved_cancel_wins_before_route_failure(
        self,
        closed_loop,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Reserved cancellation is authoritative before early Runtime work."""
        controller, _plugin, task_mgr, _ws_mgr, adapter = closed_loop
        input_file = tmp_path / "cancel-before-route.md"
        input_file.write_text("# Cancel before route", encoding="utf-8")
        intermediate = tmp_path / "preconversion-intermediate.docx"
        intermediate.write_text("must not be published", encoding="utf-8")
        output_dir = tmp_path / "cancel-before-route-out"
        request = ConversionRequest(
            request_id="cancel-before-route",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            options={
                PRECONVERSION_INTERMEDIATES_OPTION: [
                    {
                        "staging_path": str(intermediate),
                        "suggested_name": "preconversion-intermediate.docx",
                        "applies_to_input_path": str(input_file),
                    }
                ]
            },
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )
        adapter_entered = threading.Event()
        release_adapter = threading.Event()
        route_called = threading.Event()
        real_execute = adapter.execute
        results: list[ConversionResult] = []

        def paused_execute(runtime_request: ConversionRequest) -> ConversionResult:
            adapter_entered.set()
            assert release_adapter.wait(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            return real_execute(runtime_request)

        def fail_route(*_args: object, **_kwargs: object) -> tuple[str, object]:
            route_called.set()
            raise RuntimeError("route should not run after reserved cancellation")

        monkeypatch.setattr(adapter, "execute", paused_execute)
        monkeypatch.setattr(task_mgr._resolver, "resolve", fail_route)
        worker = threading.Thread(target=lambda: results.append(controller.execute_single(request)))
        worker.start()
        try:
            assert adapter_entered.wait(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            controller.cancel(request.request_id)
            release_adapter.set()
            worker.join(_THREAD_COORDINATION_TIMEOUT_SECONDS)
        finally:
            release_adapter.set()
            worker.join(_THREAD_COORDINATION_TIMEOUT_SECONDS)

        assert not worker.is_alive()
        assert route_called.is_set() is False
        assert len(results) == 1
        assert results[0].success is False
        assert results[0].error is not None
        assert results[0].error.error_type == "cancelled"
        assert results[0].artifacts == []
        assert intermediate.exists()
        assert output_dir.exists() is False
        assert task_mgr._pending_cancellations == set()
        assert task_mgr._tokens == {}
        assert task_mgr._reserved_cancellations == set()

    def test_inflight_cancel_defers_reservation_release_until_port_returns(
        self,
        closed_loop,
        tmp_path,
        monkeypatch,
    ) -> None:
        """Application retains the Runtime token while cancel is in flight."""
        controller, plugin, task_mgr, _ws_mgr, adapter = closed_loop
        input_file = tmp_path / "inflight-cancel.md"
        input_file.write_text("# Inflight cancel", encoding="utf-8")
        request = ConversionRequest(
            request_id="inflight-cancel",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "inflight-cancel-out")),
        )
        plugin_entered = threading.Event()
        release_plugin = threading.Event()
        cancel_entered = threading.Event()
        release_cancel = threading.Event()
        real_convert = plugin.convert
        real_cancel = adapter.cancel
        results: list[ConversionResult] = []
        cancel_failures: list[BaseException] = []

        def paused_convert(context: Any) -> ConversionResult:
            plugin_entered.set()
            assert release_plugin.wait(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            return real_convert(context)

        def paused_cancel(task_id: str) -> None:
            cancel_entered.set()
            assert release_cancel.wait(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            real_cancel(task_id)

        def cancel() -> None:
            try:
                controller.cancel(request.request_id)
            except BaseException as exc:  # pragma: no cover - asserted below
                cancel_failures.append(exc)

        monkeypatch.setattr(plugin, "convert", paused_convert)
        monkeypatch.setattr(adapter, "cancel", paused_cancel)
        worker = threading.Thread(target=lambda: results.append(controller.execute_single(request)))
        worker.start()
        canceller: threading.Thread | None = None
        try:
            assert plugin_entered.wait(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            canceller = threading.Thread(target=cancel)
            canceller.start()
            assert cancel_entered.wait(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            release_plugin.set()
            worker.join(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            assert not worker.is_alive()
            assert task_mgr._reserved_cancellations == {request.request_id}
            assert request.request_id in task_mgr._tokens

            release_cancel.set()
            canceller.join(_THREAD_COORDINATION_TIMEOUT_SECONDS)
        finally:
            release_plugin.set()
            release_cancel.set()
            worker.join(_THREAD_COORDINATION_TIMEOUT_SECONDS)
            if canceller is not None:
                canceller.join(_THREAD_COORDINATION_TIMEOUT_SECONDS)

        assert canceller is not None
        assert not canceller.is_alive()
        assert cancel_failures == []
        assert len(results) == 1
        assert results[0].success is True
        assert task_mgr._pending_cancellations == set()
        assert task_mgr._tokens == {}
        assert task_mgr._reserved_cancellations == set()
