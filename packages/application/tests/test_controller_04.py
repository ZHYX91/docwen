"""Focused tests split from test_controller.py."""

from __future__ import annotations

from ._controller_support import (
    ApplicationController,
    ControllerError,
    ConversionRequest,
    MagicMock,
    Path,
    _file_ref,
    _request,
    call,
    patch,
    pytest,
    shutil,
    tempfile,
    threading,
)
from ._controller_support import (
    _isolate_controller_tests_from_core_admission as _isolate_controller_tests_from_core_admission,
)
from ._controller_support import (
    mock_config as mock_config,
)
from ._controller_support import (
    mock_presenter as mock_presenter,
)
from ._controller_support import (
    mock_runtime as mock_runtime,
)

pytestmark = pytest.mark.unit


class TestExecuteMethods:
    """execute_single / execute_batch delegate correctly."""

    def test_admitted_rtf_with_disguised_suffix_is_bridged_from_canonical_copy(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_core.models.result import ConversionResult
        from docwen_core.office_bridge import BridgeResult

        source = tmp_path / "notice.bin"
        source_bytes = b"{\\rtf1\\ansi Protective copy probe}"
        source.write_bytes(source_bytes)
        protected_inputs: list[tuple[Path, bytes]] = []

        def fake_bridge(input_path, output_path, **_kwargs):
            protected = Path(input_path)
            protected_inputs.append((protected, protected.read_bytes()))
            Path(output_path).write_bytes(b"converted docx")
            return BridgeResult(True, output_path=output_path, backend="Fake Office")

        def fake_runtime(child: ConversionRequest) -> ConversionResult:
            assert Path(child.input_refs[0].path).name == "notice.docx"
            return ConversionResult(task_id=child.request_id, success=True)

        mock_runtime.execute.side_effect = fake_runtime
        ctrl = ApplicationController(runtime_port=mock_runtime)
        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        try:
            with patch(
                "docwen_application.preconversion.pre_converter.convert_with_backend_priority",
                side_effect=fake_bridge,
            ):
                result = ctrl.execute_single(
                    ConversionRequest(
                        request_id="admitted-rtf",
                        input_refs=[_file_ref(str(source), "rtf")],
                        target_format="md",
                    )
                )

            assert result.success is True
            assert [(path.name, content) for path, content in protected_inputs] == [("input.rtf", source_bytes)]
            assert source.read_bytes() == source_bytes
            assert all(not path.exists() for path, _content in protected_inputs)
            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_cancel_before_execution_stops_preconversion_without_runtime_pending(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        source = tmp_path / "legacy.doc"
        source.write_text("legacy", encoding="utf-8")
        request = _request("cancel-before-start", str(source), source_format="doc")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))

        reservation = ctrl.prepare_execution_cancellation(request)
        ctrl.cancel(request.request_id)
        try:
            with (
                patch(
                    "docwen_application.preconversion.chain_resolver.resolve_chain",
                    return_value=["docx", "md"],
                ),
                patch("docwen_application.preconversion.pre_converter.pre_convert") as pre_convert_mock,
            ):
                result = ctrl.execute_single(request)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "cancelled"
            pre_convert_mock.assert_not_called()
            mock_runtime.execute.assert_not_called()
            mock_runtime.cancel.assert_not_called()
            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            ctrl.release_execution_cancellation(request.request_id, reservation)
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_cancel_during_preconversion_reaches_bridge_without_runtime_admission(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "legacy.doc"
        source.write_text("legacy", encoding="utf-8")
        request = _request("cancel-preconversion", str(source), source_format="doc")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        bridge_entered = threading.Event()
        release_bridge = threading.Event()
        bridge_saw_cancel = threading.Event()
        results: list[ConversionResult] = []
        failures: list[BaseException] = []

        def fake_pre_convert(
            input_path: str,
            _source_format: str,
            *,
            staging_dir: str,
            cancel: object,
            **_kwargs: object,
        ) -> PreConversionResult:
            output = Path(staging_dir) / "legacy.docx"
            output.write_text(Path(input_path).read_text(encoding="utf-8"), encoding="utf-8")
            bridge_entered.set()
            assert release_bridge.wait(2.0)
            if bool(getattr(cancel, "is_cancelled", False)):
                bridge_saw_cancel.set()
            return PreConversionResult(str(output), "doc", "Fake Office")

        def execute() -> None:
            try:
                results.append(ctrl.execute_single(request))
            except BaseException as exc:  # pragma: no cover - asserted below
                failures.append(exc)

        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        try:
            with (
                patch(
                    "docwen_application.preconversion.chain_resolver.resolve_chain",
                    return_value=["docx", "md"],
                ),
                patch(
                    "docwen_application.preconversion.pre_converter.pre_convert",
                    side_effect=fake_pre_convert,
                ),
            ):
                worker = threading.Thread(target=execute)
                worker.start()
                assert bridge_entered.wait(2.0)
                ctrl.cancel(request.request_id)
                release_bridge.set()
                worker.join(2.0)

            assert not worker.is_alive()
            assert failures == []
            assert bridge_saw_cancel.is_set()
            assert len(results) == 1
            result = results[0]
            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "cancelled"
            mock_runtime.execute.assert_not_called()
            mock_runtime.cancel.assert_not_called()
            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            release_bridge.set()
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_bridge_cancelled_outcome_projects_structured_cancellation(
        self,
        mock_runtime: MagicMock,
        tmp_path,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult

        source = tmp_path / "legacy.doc"
        source.write_text("legacy", encoding="utf-8")
        ctrl = ApplicationController(runtime_port=mock_runtime)

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=["docx", "md"],
            ),
            patch(
                "docwen_application.preconversion.pre_converter.convert_with_backend_priority",
                return_value=BridgeResult(
                    False,
                    message="bridge stopped",
                    cancelled=True,
                    error_code="OFFICE_CONVERSION_CANCELLED",
                    cleanup_message="Private Office workspace cleanup failed: test workspace",
                    cleanup_failed=True,
                ),
            ),
        ):
            result = ctrl.execute_single(_request("bridge-cancelled", str(source), source_format="doc"))

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "cancelled"
        assert result.error.message == "Task was cancelled"
        assert [(item.level, item.code, item.message) for item in result.diagnostics] == [
            (
                "warning",
                "OFFICE_CLEANUP_FAILED",
                "Private Office workspace cleanup failed: test workspace",
            )
        ]
        mock_runtime.execute.assert_not_called()
        mock_runtime.cancel.assert_not_called()

    def test_preconversion_batch_cancel_preserves_prior_failure_and_cancels_uncommitted(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        sources = [tmp_path / name for name in ("first.doc", "second.doc", "third.doc")]
        for source in sources:
            source.write_text(source.stem, encoding="utf-8")
        request = _request(
            "cancel-mixed-batch",
            *(str(source) for source in sources),
            source_format="doc",
        )
        ctrl = ApplicationController(runtime_port=mock_runtime)
        second_entered = threading.Event()
        release_second = threading.Event()
        calls: list[str] = []
        results: list[list[ConversionResult]] = []
        failures: list[BaseException] = []

        def fake_pre_convert(
            input_path: str,
            _source_format: str,
            *,
            staging_dir: str,
            **_kwargs: object,
        ) -> PreConversionResult | None:
            calls.append(Path(input_path).name)
            if Path(input_path).name == "first.doc":
                return None
            output = Path(staging_dir) / f"{Path(input_path).stem}.docx"
            output.write_text("hub", encoding="utf-8")
            if Path(input_path).name == "second.doc":
                second_entered.set()
                assert release_second.wait(2.0)
            return PreConversionResult(str(output), "doc", "Fake Office")

        def execute() -> None:
            try:
                results.append(ctrl.execute_batch(request))
            except BaseException as exc:  # pragma: no cover - asserted below
                failures.append(exc)

        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        try:
            with (
                patch(
                    "docwen_application.preconversion.chain_resolver.resolve_chain",
                    return_value=["docx", "md"],
                ),
                patch(
                    "docwen_application.preconversion.pre_converter.pre_convert",
                    side_effect=fake_pre_convert,
                ),
            ):
                worker = threading.Thread(target=execute)
                worker.start()
                assert second_entered.wait(2.0)
                ctrl.cancel(request.request_id)
                release_second.set()
                worker.join(2.0)

            assert not worker.is_alive()
            assert failures == []
            assert calls == ["first.doc", "second.doc"]
            assert len(results) == 1
            assert [result.task_id for result in results[0]] == [
                "cancel-mixed-batch-0",
                "cancel-mixed-batch-1",
                "cancel-mixed-batch-2",
            ]
            assert [result.error.error_type if result.error is not None else "" for result in results[0]] == [
                "dependency_missing",
                "cancelled",
                "cancelled",
            ]
            mock_runtime.execute.assert_not_called()
            mock_runtime.cancel.assert_not_called()
            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            release_second.set()
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_cancel_active_batch_parent_forwards_only_active_child_once(
        self,
        mock_runtime: MagicMock,
    ) -> None:
        from docwen_core.models.result import ConversionResult

        request = _request("cancel-batch", "/first.md", "/second.md")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        first_started = threading.Event()
        release_first = threading.Event()
        results: list[list[ConversionResult]] = []
        failures: list[BaseException] = []

        def fake_execute(child_request: ConversionRequest) -> ConversionResult:
            if child_request.request_id == "cancel-batch-0":
                first_started.set()
                assert release_first.wait(2.0)
            return ConversionResult(task_id=child_request.request_id, success=True)

        def execute() -> None:
            try:
                results.append(ctrl.execute_batch(request))
            except BaseException as exc:  # pragma: no cover - asserted below
                failures.append(exc)

        mock_runtime.execute.side_effect = fake_execute
        worker = threading.Thread(target=execute)
        worker.start()
        try:
            assert first_started.wait(2.0)
            ctrl.cancel(request.request_id)
            ctrl.cancel(request.request_id)
            release_first.set()
            worker.join(2.0)

            assert not worker.is_alive()
            assert failures == []
            assert len(results) == 1
            assert mock_runtime.cancel.call_args_list == [call("cancel-batch-0")]
            assert [result.error.error_type if result.error is not None else "" for result in results[0]] == [
                "",
                "cancelled",
            ]
        finally:
            release_first.set()

    def test_cancel_mid_batch_does_not_forward_terminal_or_future_children(
        self,
        mock_runtime: MagicMock,
    ) -> None:
        from docwen_core.models.result import ConversionResult

        request = _request("cancel-mid-batch", "/first.md", "/second.md", "/third.md")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        second_started = threading.Event()
        release_second = threading.Event()
        active: set[str] = set()
        pending: set[str] = set()
        results: list[list[ConversionResult]] = []

        def fake_execute(child_request: ConversionRequest) -> ConversionResult:
            task_id = child_request.request_id
            active.add(task_id)
            try:
                if task_id == "cancel-mid-batch-1":
                    second_started.set()
                    assert release_second.wait(2.0)
                return ConversionResult(task_id=task_id, success=True)
            finally:
                active.discard(task_id)

        def fake_cancel(task_id: str) -> None:
            if task_id not in active:
                pending.add(task_id)

        mock_runtime.execute.side_effect = fake_execute
        mock_runtime.cancel.side_effect = fake_cancel

        worker = threading.Thread(target=lambda: results.append(ctrl.execute_batch(request)))
        worker.start()
        try:
            assert second_started.wait(2.0)
            ctrl.cancel(request.request_id)
            release_second.set()
            worker.join(2.0)

            assert not worker.is_alive()
            assert mock_runtime.cancel.call_args_list == [call("cancel-mid-batch-1")]
            assert pending == set()
            assert len(results) == 1
            assert [result.task_id for result in results[0]] == [
                "cancel-mid-batch-0",
                "cancel-mid-batch-1",
                "cancel-mid-batch-2",
            ]
            assert [result.error.error_type if result.error is not None else "" for result in results[0]] == [
                "",
                "",
                "cancelled",
            ]
        finally:
            release_second.set()

    def test_active_runtime_cancel_failure_can_be_retried(
        self,
        mock_runtime: MagicMock,
    ) -> None:
        from docwen_core.models.result import ConversionResult

        request = _request("cancel-retry", "/input.md")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        runtime_started = threading.Event()
        release_runtime = threading.Event()

        def fake_execute(child_request: ConversionRequest) -> ConversionResult:
            runtime_started.set()
            assert release_runtime.wait(2.0)
            return ConversionResult(task_id=child_request.request_id, success=True)

        mock_runtime.execute.side_effect = fake_execute
        mock_runtime.cancel.side_effect = [OSError("transient cancel failure"), None]
        worker = threading.Thread(target=lambda: ctrl.execute_single(request))
        worker.start()
        try:
            assert runtime_started.wait(2.0)
            with pytest.raises(OSError, match="transient cancel failure"):
                ctrl.cancel(request.request_id)
            ctrl.cancel(request.request_id)
            assert mock_runtime.cancel.call_args_list == [call("cancel-retry"), call("cancel-retry")]
        finally:
            release_runtime.set()
            worker.join(2.0)
        assert not worker.is_alive()

    def test_command_construction_failure_never_forwards_runtime_cancel(
        self,
        mock_runtime: MagicMock,
    ) -> None:
        request = _request("command-construction", "/input.md")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        construction_started = threading.Event()
        release_construction = threading.Event()
        failures: list[BaseException] = []

        def fail_command_construction() -> object:
            construction_started.set()
            assert release_construction.wait(2.0)
            raise RuntimeError("command construction failed")

        def execute() -> None:
            try:
                ctrl.execute_single(request)
            except BaseException as exc:  # pragma: no cover - asserted below
                failures.append(exc)

        with patch.object(ctrl, "_convert_command", side_effect=fail_command_construction):
            worker = threading.Thread(target=execute)
            worker.start()
            assert construction_started.wait(2.0)
            ctrl.cancel(request.request_id)
            release_construction.set()
            worker.join(2.0)

        assert not worker.is_alive()
        assert len(failures) == 1
        assert str(failures[0]) == "command construction failed"
        mock_runtime.execute.assert_not_called()
        mock_runtime.cancel.assert_not_called()

    def test_unknown_unreserved_cancel_never_reaches_runtime(
        self,
        mock_runtime: MagicMock,
    ) -> None:
        ctrl = ApplicationController(runtime_port=mock_runtime)

        ctrl.cancel("aggregate-parent")

        mock_runtime.cancel.assert_not_called()

    def test_retained_scope_requires_identity_release_before_task_id_reuse(
        self,
        mock_runtime: MagicMock,
    ) -> None:
        from docwen_core.models.result import ConversionResult

        request = _request("retained-generation", "/input.md")
        mock_runtime.execute.return_value = ConversionResult(task_id=request.request_id, success=True)
        ctrl = ApplicationController(runtime_port=mock_runtime)

        first_reservation = ctrl.prepare_execution_cancellation(request)
        assert ctrl.execute_single(request).success is True
        with pytest.raises(ControllerError, match="awaiting release"):
            ctrl.prepare_execution_cancellation(request)

        ctrl.release_execution_cancellation(request.request_id, first_reservation)
        second_reservation = ctrl.prepare_execution_cancellation(request)
        ctrl.release_execution_cancellation(request.request_id, first_reservation)

        assert ctrl._cancellation_scopes[request.request_id] is second_reservation
        ctrl.release_execution_cancellation(request.request_id, second_reservation)
        assert ctrl._cancellation_scopes == {}
