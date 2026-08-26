"""Focused tests split from test_controller.py."""

from __future__ import annotations

from ._controller_support import (
    Any,
    ApplicationController,
    ConversionRequest,
    FileRef,
    MagicMock,
    OutputPolicy,
    Path,
    _request,
    patch,
    pytest,
    shutil,
    tempfile,
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

    @pytest.mark.parametrize(
        ("actual_format", "target_format", "path"),
        [
            ("doc", "docx", "/legacy.doc"),
            ("xls", "xlsx", "/legacy.xls"),
        ],
    )
    def test_execute_single_hub_target_uses_direct_plugin_route(
        self,
        mock_runtime: MagicMock,
        actual_format: str,
        target_format: str,
        path: str,
    ) -> None:
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id=f"{actual_format}-{target_format}", success=True)
        mock_runtime.execute.return_value = expected

        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = _request(
            expected.task_id,
            path,
            target_format=target_format,
            source_format=actual_format,
        )

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=[target_format],
            ),
            patch("docwen_application.preconversion.pre_converter.pre_convert") as pre_convert_mock,
        ):
            result = ctrl.execute_single(request)

        assert result is expected
        pre_convert_mock.assert_not_called()
        runtime_request = mock_runtime.execute.call_args[0][0]
        assert runtime_request.input_refs[0].path == path
        assert runtime_request.target_format == target_format

    def test_execute_single_xls_to_md_uses_spreadsheet_plugin_bridge(self, mock_runtime: MagicMock) -> None:
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="xls-md", success=True)
        mock_runtime.execute.return_value = expected

        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = _request("xls-md", "/legacy.xls", target_format="md", source_format="xls")

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=["md"],
            ),
            patch("docwen_application.preconversion.pre_converter.pre_convert") as pre_convert_mock,
        ):
            result = ctrl.execute_single(request)

        assert result is expected
        pre_convert_mock.assert_not_called()
        runtime_request = mock_runtime.execute.call_args[0][0]
        assert runtime_request.input_refs[0].path == "/legacy.xls"
        assert runtime_request.target_format == "md"

    @pytest.mark.parametrize("output_mode", ["source-none", "source-empty", "custom"])
    def test_preconversion_batch_isolates_same_stem_preserves_output_policy_and_cleans_staging(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
        output_mode: str,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        first_dir = tmp_path / "first"
        second_dir = tmp_path / "second"
        first_dir.mkdir()
        second_dir.mkdir()
        first = first_dir / "report.doc"
        second = second_dir / "report.doc"
        first.write_text("FIRST", encoding="utf-8")
        second.write_text("SECOND", encoding="utf-8")
        configured_output_dir = {
            "source-none": None,
            "source-empty": "",
            "custom": str(tmp_path / "custom-output"),
        }[output_mode]
        expected_output_dirs = (
            [configured_output_dir, configured_output_dir]
            if configured_output_dir
            else [str(first_dir), str(second_dir)]
        )
        staging_dirs: list[Path] = []
        seen: list[tuple[Path, str, dict[str, object], dict[str, Any], str]] = []

        def fake_pre_convert(input_path: str, _source_format: str, *, staging_dir: str, **_kwargs):
            stage = Path(staging_dir)
            stage.mkdir(parents=True, exist_ok=True)
            staging_dirs.append(stage)
            output = stage / f"{Path(input_path).stem}.docx"
            output.write_text(Path(input_path).read_text(encoding="utf-8"), encoding="utf-8")
            return PreConversionResult(str(output), "doc", "Fake Office")

        def fake_runtime_execute(runtime_request: ConversionRequest) -> ConversionResult:
            runtime_ref = runtime_request.input_refs[0]
            runtime_path = Path(runtime_ref.path)
            seen.append(
                (
                    runtime_path,
                    runtime_path.read_text(encoding="utf-8"),
                    runtime_request.output_policy.to_dict(),
                    dict(runtime_ref.metadata),
                    runtime_ref.warning_message,
                )
            )
            return ConversionResult(task_id=runtime_request.request_id, success=True)

        mock_runtime.execute.side_effect = fake_runtime_execute
        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = ConversionRequest(
            request_id="same-stem",
            input_refs=[
                FileRef(
                    path=str(first),
                    format="doc",
                    category="document",
                    warning_message="first warning",
                    metadata={"detector": "first"},
                ),
                FileRef(
                    path=str(second),
                    format="doc",
                    category="document",
                    warning_message="second warning",
                    metadata={"detector": "second"},
                ),
            ],
            target_format="md",
            output_policy=OutputPolicy(
                output_dir=configured_output_dir,
                date_subfolder="compact",
                overwrite_mode="skip",
                open_after_done=True,
            ),
        )

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
                results = ctrl.execute_batch(request)

            assert [result.success for result in results] == [True, True]
            assert [item[1] for item in seen] == ["FIRST", "SECOND"]
            assert seen[0][0] != seen[1][0]
            assert [item[0].name for item in seen] == ["report.docx", "report.docx"]
            assert [item[2] for item in seen] == [
                {
                    "output_dir": expected_output_dirs[0],
                    "output_path": None,
                    "date_subfolder": "compact",
                    "overwrite_mode": "skip",
                    "write_artifacts": True,
                    "open_after_done": True,
                },
                {
                    "output_dir": expected_output_dirs[1],
                    "output_path": None,
                    "date_subfolder": "compact",
                    "overwrite_mode": "skip",
                    "write_artifacts": True,
                    "open_after_done": True,
                },
            ]
            assert [item[3]["detector"] for item in seen] == ["first", "second"]
            provenance = [item[3]["_docwen_preconversion_source"] for item in seen]
            assert [item["path"] for item in provenance] == [str(first), str(second)]
            assert [item["format"] for item in provenance] == ["doc", "doc"]
            assert [item["category"] for item in provenance] == ["document", "document"]
            assert [item["warning_message"] for item in provenance] == ["first warning", "second warning"]
            assert [item["inspection"] for item in provenance] == [None, None]
            assert [item[4] for item in seen] == ["", ""]
            assert [result.task_id for result in results] == ["same-stem-0", "same-stem-1"]
            assert staging_dirs and not list(tmp_path.glob("docwen_pre_*"))
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_preconversion_staging_is_cleaned_when_runtime_raises(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult

        source = tmp_path / "legacy.doc"
        source.write_text("legacy", encoding="utf-8")

        def fake_pre_convert(input_path: str, _source_format: str, *, staging_dir: str, **_kwargs):
            stage = Path(staging_dir)
            stage.mkdir(parents=True, exist_ok=True)
            output = stage / "legacy.docx"
            output.write_text(Path(input_path).read_text(encoding="utf-8"), encoding="utf-8")
            return PreConversionResult(str(output), "doc", "Fake Office")

        mock_runtime.execute.side_effect = RuntimeError("runtime exploded")
        ctrl = ApplicationController(runtime_port=mock_runtime)
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
                pytest.raises(RuntimeError, match="runtime exploded"),
            ):
                ctrl.execute_single(_request("runtime-error", str(source), source_format="doc"))

            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_preconversion_staging_is_cleaned_when_backend_is_unavailable(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        class UnexpectedDeepcopy:
            def __deepcopy__(self, _memo):
                raise AssertionError("failed preconversion must not rebuild options")

        source = tmp_path / "legacy.doc"
        source.write_text("legacy", encoding="utf-8")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = _request("backend-unavailable", str(source), source_format="doc")
        request.options["opaque"] = UnexpectedDeepcopy()
        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        try:
            with (
                patch(
                    "docwen_application.preconversion.chain_resolver.resolve_chain",
                    return_value=["docx", "md"],
                ),
                patch(
                    "docwen_application.preconversion.pre_converter.pre_convert",
                    return_value=None,
                ),
            ):
                result = ctrl.execute_single(request)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "dependency_missing"
            assert not list(tmp_path.glob("docwen_pre_*"))
            mock_runtime.execute.assert_not_called()
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_preconversion_copy_failure_is_structured_and_cleans_staging(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult

        source = tmp_path / "locked.disguised"
        source.write_bytes(b"legacy")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        try:
            with (
                patch(
                    "docwen_application.preconversion.chain_resolver.resolve_chain",
                    return_value=["docx", "md"],
                ),
                patch(
                    "docwen_application.preconversion.pre_converter._copy_snapshot_stream",
                    side_effect=PermissionError("source is locked"),
                ) as copy_mock,
                patch(
                    "docwen_application.preconversion.pre_converter.convert_with_backend_priority",
                    return_value=BridgeResult(False, message="bridge must not run"),
                ) as bridge_mock,
            ):
                result = ctrl.execute_single(_request("copy-failure", str(source), source_format="doc"))

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"
            assert result.error.diagnostic_code == "PRECONVERSION_INPUT_COPY_FAILED"
            assert "protective input copy" in result.error.message
            copy_mock.assert_called_once()
            bridge_mock.assert_not_called()
            mock_runtime.execute.assert_not_called()
            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_installed_preconversion_backend_failure_is_not_reported_as_missing_dependency(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult

        source = tmp_path / "legacy.doc"
        source.write_bytes(b"legacy")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        try:
            with (
                patch(
                    "docwen_application.preconversion.chain_resolver.resolve_chain",
                    return_value=["docx", "md"],
                ),
                patch(
                    "docwen_application.preconversion.pre_converter.convert_with_backend_priority",
                    return_value=BridgeResult(
                        False,
                        message="Microsoft Word export failed",
                        attempted_backend_ids=("msoffice_word",),
                        available_backend_ids=("msoffice_word",),
                        error_code="OFFICE_BACKEND_FAILED",
                        cleanup_message="Private Office workspace cleanup failed: test workspace",
                        cleanup_failed=True,
                    ),
                ),
            ):
                result = ctrl.execute_single(_request("installed-backend-failure", str(source), source_format="doc"))

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"
            assert result.error.diagnostic_code == "OFFICE_BACKEND_FAILED"
            assert "Microsoft Word export failed" in result.error.message
            assert [(item.level, item.code, item.message) for item in result.diagnostics] == [
                (
                    "warning",
                    "OFFICE_CLEANUP_FAILED",
                    "Private Office workspace cleanup failed: test workspace",
                )
            ]
            mock_runtime.execute.assert_not_called()
            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_cancel_after_protective_copy_skips_bridge_and_cleans_staging(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        source = tmp_path / "legacy.doc"
        source.write_bytes(b"legacy")
        request = _request("cancel-after-copy", str(source), source_format="doc")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        from docwen_application.preconversion import pre_converter

        real_copy = pre_converter._copy_snapshot_stream

        def copy_then_cancel(source_stream, destination_stream, token):
            copied = real_copy(source_stream, destination_stream, token)
            ctrl.cancel(request.request_id)
            return copied

        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        try:
            with (
                patch(
                    "docwen_application.preconversion.chain_resolver.resolve_chain",
                    return_value=["docx", "md"],
                ),
                patch(
                    "docwen_application.preconversion.pre_converter._copy_snapshot_stream",
                    side_effect=copy_then_cancel,
                ) as copy_mock,
                patch("docwen_application.preconversion.pre_converter.convert_with_backend_priority") as bridge_mock,
            ):
                result = ctrl.execute_single(request)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "cancelled"
            copy_mock.assert_called_once()
            bridge_mock.assert_not_called()
            mock_runtime.execute.assert_not_called()
            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_preconversion_copy_failure_stays_aligned_in_batch(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_core.models.result import ConversionResult
        from docwen_core.office_bridge import BridgeResult

        first = tmp_path / "first.disguised"
        second = tmp_path / "second.disguised"
        first.write_bytes(b"FIRST")
        second.write_bytes(b"SECOND")
        request = _request("copy-batch", str(first), str(second), source_format="doc")
        ctrl = ApplicationController(runtime_port=mock_runtime)
        from docwen_application.preconversion import pre_converter

        real_copy = pre_converter._copy_snapshot_stream
        protected_inputs: list[tuple[Path, bytes]] = []
        copy_calls = 0

        def selective_copy(source_stream, destination_stream, token):
            nonlocal copy_calls
            copy_calls += 1
            if copy_calls == 1:
                raise PermissionError("first source is locked")
            return real_copy(source_stream, destination_stream, token)

        def fake_bridge(input_path, output_path, **_kwargs):
            protected = Path(input_path)
            protected_inputs.append((protected, protected.read_bytes()))
            Path(output_path).write_bytes(protected.read_bytes())
            return BridgeResult(True, output_path=output_path, backend="Fake Office")

        mock_runtime.execute.side_effect = lambda child: ConversionResult(task_id=child.request_id, success=True)
        monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
        try:
            with (
                patch(
                    "docwen_application.preconversion.chain_resolver.resolve_chain",
                    return_value=["docx", "md"],
                ),
                patch(
                    "docwen_application.preconversion.pre_converter._copy_snapshot_stream",
                    side_effect=selective_copy,
                ),
                patch(
                    "docwen_application.preconversion.pre_converter.convert_with_backend_priority",
                    side_effect=fake_bridge,
                ),
            ):
                results = ctrl.execute_batch(request)

            assert [result.task_id for result in results] == ["copy-batch-0", "copy-batch-1"]
            assert [result.success for result in results] == [False, True]
            assert results[0].error is not None
            assert results[0].error.error_type == "conversion_failed"
            assert results[0].error.diagnostic_code == "PRECONVERSION_INPUT_COPY_FAILED"
            assert [(path.name, content) for path, content in protected_inputs] == [("input.doc", b"SECOND")]
            assert mock_runtime.execute.call_count == 1
            runtime_request = mock_runtime.execute.call_args[0][0]
            assert runtime_request.request_id == "copy-batch-1"
            assert Path(runtime_request.input_refs[0].path).name == "second.docx"
            assert not list(tmp_path.glob("docwen_pre_*"))
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)
