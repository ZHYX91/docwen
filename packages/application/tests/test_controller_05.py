"""Focused tests split from test_controller.py."""

from __future__ import annotations

from ._controller_support import (
    PRECONVERSION_INTERMEDIATES_OPTION,
    ApplicationController,
    ControllerError,
    ConversionRequest,
    MagicMock,
    OutputPolicy,
    Path,
    _file_ref,
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

    def test_cleanup_failure_still_completes_direct_cancellation_scope(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "legacy.doc"
        source.write_text("legacy", encoding="utf-8")
        request = _request("cleanup-scope", str(source), source_format="doc")
        mock_runtime.execute.return_value = ConversionResult(task_id=request.request_id, success=True)
        ctrl = ApplicationController(runtime_port=mock_runtime)

        def fake_pre_convert(input_path: str, _source_format: str, *, staging_dir: str, **_kwargs: object):
            output = Path(staging_dir) / "legacy.docx"
            output.write_text(Path(input_path).read_text(encoding="utf-8"), encoding="utf-8")
            return PreConversionResult(str(output), "doc", "Fake Office")

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
                patch(
                    "docwen_application.controller._ManagedPreconversion.cleanup",
                    side_effect=RuntimeError("cleanup failed"),
                ),
                pytest.raises(RuntimeError, match="cleanup failed"),
            ):
                ctrl.execute_single(request)

            assert ctrl._cancellation_scopes == {}
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_preconversion_staging_is_cleaned_when_backend_raises(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        source = tmp_path / "legacy.doc"
        source.write_text("legacy", encoding="utf-8")
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
                    side_effect=RuntimeError("bridge exploded"),
                ),
                pytest.raises(RuntimeError, match="bridge exploded"),
            ):
                ctrl.execute_single(_request("backend-error", str(source), source_format="doc"))

            assert not list(tmp_path.glob("docwen_pre_*"))
            mock_runtime.execute.assert_not_called()
        finally:
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    def test_preconversion_staging_is_cleaned_when_request_rebuild_raises(
        self,
        mock_runtime: MagicMock,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult

        class DeepcopyBomb:
            def __deepcopy__(self, _memo):
                raise RuntimeError("options deepcopy exploded")

        real_temporary_directory = tempfile.TemporaryDirectory
        owners = []

        class TrackingTemporaryDirectory:
            def __init__(self, **kwargs):
                self._inner = real_temporary_directory(**kwargs)
                self.name = self._inner.name
                self.cleanup_called = False
                owners.append(self)

            def cleanup(self) -> None:
                self.cleanup_called = True
                self._inner.cleanup()

        source = tmp_path / "legacy.doc"
        source.write_text("legacy", encoding="utf-8")

        def fake_pre_convert(input_path: str, _source_format: str, *, staging_dir: str, **_kwargs):
            stage = Path(staging_dir)
            stage.mkdir(parents=True, exist_ok=True)
            output = stage / "legacy.docx"
            output.write_text(Path(input_path).read_text(encoding="utf-8"), encoding="utf-8")
            return PreConversionResult(str(output), "doc", "Fake Office")

        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = ConversionRequest(
            request_id="request-rebuild-error",
            input_refs=[_file_ref(str(source), "doc")],
            target_format="md",
            options={"bomb": DeepcopyBomb()},
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
                patch(
                    "docwen_application.controller.tempfile.TemporaryDirectory",
                    side_effect=TrackingTemporaryDirectory,
                ),
                pytest.raises(RuntimeError, match="options deepcopy exploded"),
            ):
                ctrl.execute_single(request)

            assert len(owners) == 1
            assert owners[0].cleanup_called is True
            assert not list(tmp_path.glob("docwen_pre_*"))
            mock_runtime.execute.assert_not_called()
        finally:
            for owner in owners:
                owner.cleanup()
            for root in tmp_path.glob("docwen_pre_*"):
                shutil.rmtree(root, ignore_errors=True)

    @pytest.mark.parametrize(
        ("actual_format", "target_format", "path"),
        [
            ("doc", "odt", "/legacy.doc"),
            ("rtf", "pdf", "/legacy.rtf"),
        ],
    )
    def test_execute_single_document_non_markdown_target_uses_plugin_route(
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

    def test_execute_single_raises_without_runtime(self) -> None:
        ctrl = ApplicationController()
        request = _request("r1", "/f.txt")
        with pytest.raises(ControllerError, match="No runtime port"):
            ctrl.execute_single(request)

    def test_execute_batch_raises_without_runtime(self) -> None:
        ctrl = ApplicationController()
        request = _request("r1", "/f.txt")
        with pytest.raises(ControllerError, match="No runtime port"):
            ctrl.execute_batch(request)

    def test_preconversion_records_intermediate_when_setting_enabled(
        self,
        mock_runtime: MagicMock,
        tmp_path,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "legacy.doc"
        source.write_text("legacy")
        preconverted = tmp_path / "legacy_pre.docx"
        preconverted.write_text("hub")
        mock_runtime.execute.return_value = ConversionResult(task_id="legacy-md", success=True)

        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = ConversionRequest(
            request_id="legacy-md",
            input_refs=[_file_ref(str(source), "doc")],
            target_format="md",
            output_policy=OutputPolicy(
                output_dir=str(tmp_path / "out"),
                date_subfolder="iso",
                overwrite_mode="overwrite",
                open_after_done=True,
            ),
            config_snapshot={"output": {"intermediate_files": {"save_to_output": True}}},
        )

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=["docx", "md"],
            ),
            patch(
                "docwen_application.preconversion.pre_converter.pre_convert",
                return_value=PreConversionResult(
                    pre_converted_path=str(preconverted),
                    original_source_format="doc",
                    backend="Fake Office",
                ),
            ),
        ):
            ctrl.execute_single(request)

        runtime_request = mock_runtime.execute.call_args[0][0]
        records = runtime_request.options[PRECONVERSION_INTERMEDIATES_OPTION]
        assert records == [
            {
                "staging_path": str(preconverted),
                "suggested_name": "legacy_fromDoc.docx",
                "source_format": "doc",
                "target_format": "docx",
                "backend": "Fake Office",
                "applies_to_input_path": str(preconverted),
            }
        ]
        assert runtime_request.input_refs[0].path == str(preconverted)
        assert runtime_request.output_policy == request.output_policy

    @pytest.mark.parametrize(
        "snapshot_value",
        [True, False],
    )
    def test_preconversion_intermediate_save_prefers_request_snapshot(
        self,
        mock_runtime: MagicMock,
        mock_config: MagicMock,
        tmp_path,
        *,
        snapshot_value: bool,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "legacy.doc"
        source.write_text("legacy")
        preconverted = tmp_path / "legacy_pre.docx"
        preconverted.write_text("hub")
        mock_runtime.execute.return_value = ConversionResult(task_id="snapshot-intermediate", success=True)

        ctrl = ApplicationController(runtime_port=mock_runtime, config_port=mock_config)
        request = ConversionRequest(
            request_id="snapshot-intermediate",
            input_refs=[_file_ref(str(source), "doc")],
            target_format="md",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
            config_snapshot={
                "output": {
                    "intermediate_files": {
                        "save_to_output": snapshot_value,
                    }
                }
            },
        )

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=["docx", "md"],
            ),
            patch(
                "docwen_application.preconversion.pre_converter.pre_convert",
                return_value=PreConversionResult(
                    pre_converted_path=str(preconverted),
                    original_source_format="doc",
                    backend="Fake Office",
                ),
            ) as pre_convert_mock,
        ):
            ctrl.execute_single(request)

        runtime_request = mock_runtime.execute.call_args[0][0]
        mock_config.snapshot.assert_not_called()
        mock_config.get.assert_not_called()
        assert pre_convert_mock.call_args.kwargs["backend_priority"] == [
            "wps_writer",
            "msoffice_word",
            "libreoffice",
        ]
        assert (PRECONVERSION_INTERMEDIATES_OPTION in runtime_request.options) is snapshot_value

    @pytest.mark.parametrize(
        "snapshot_value",
        [True, False],
    )
    def test_preconversion_intermediate_save_captures_config_port_snapshot(
        self,
        mock_runtime: MagicMock,
        mock_config: MagicMock,
        tmp_path,
        *,
        snapshot_value: bool,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "legacy.doc"
        source.write_text("legacy")
        preconverted = tmp_path / "legacy_pre.docx"
        preconverted.write_text("hub")
        mock_runtime.execute.return_value = ConversionResult(task_id="captured-intermediate", success=True)
        captured_snapshot = {
            "output": {
                "intermediate_files": {
                    "save_to_output": snapshot_value,
                }
            },
            "software": {
                "default_priority": {
                    "word_processors": ["libreoffice"],
                }
            },
        }
        mock_config.snapshot.return_value = captured_snapshot

        ctrl = ApplicationController(runtime_port=mock_runtime, config_port=mock_config)
        request = ConversionRequest(
            request_id="captured-intermediate",
            input_refs=[_file_ref(str(source), "doc")],
            target_format="md",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=["docx", "md"],
            ),
            patch(
                "docwen_application.preconversion.pre_converter.pre_convert",
                return_value=PreConversionResult(
                    pre_converted_path=str(preconverted),
                    original_source_format="doc",
                    backend="Fake Office",
                ),
            ) as pre_convert_mock,
        ):
            ctrl.execute_single(request)

        mock_config.snapshot.assert_called_once_with()
        mock_config.get.assert_not_called()
        assert pre_convert_mock.call_args.kwargs["backend_priority"] == ["libreoffice"]
        runtime_request = mock_runtime.execute.call_args[0][0]
        assert runtime_request.config_snapshot == captured_snapshot
        assert runtime_request.config_snapshot is not captured_snapshot
        assert runtime_request.config_snapshot["output"] is not captured_snapshot["output"]
        assert (PRECONVERSION_INTERMEDIATES_OPTION in runtime_request.options) is snapshot_value
        assert request.config_snapshot == {}
        assert request.options == {}

    def test_preconversion_empty_config_port_snapshot_is_authoritative(
        self,
        mock_runtime: MagicMock,
        mock_config: MagicMock,
        tmp_path,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "legacy.doc"
        source.write_text("legacy")
        preconverted = tmp_path / "legacy_pre.docx"
        preconverted.write_text("hub")
        mock_runtime.execute.return_value = ConversionResult(task_id="empty-captured-snapshot", success=True)
        mock_config.snapshot.return_value = {}

        ctrl = ApplicationController(runtime_port=mock_runtime, config_port=mock_config)
        request = ConversionRequest(
            request_id="empty-captured-snapshot",
            input_refs=[_file_ref(str(source), "doc")],
            target_format="md",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
        )

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=["docx", "md"],
            ),
            patch(
                "docwen_application.preconversion.pre_converter.pre_convert",
                return_value=PreConversionResult(
                    pre_converted_path=str(preconverted),
                    original_source_format="doc",
                    backend="Fake Office",
                ),
            ),
        ):
            ctrl.execute_single(request)

        mock_config.snapshot.assert_called_once_with()
        mock_config.get.assert_not_called()
        runtime_request = mock_runtime.execute.call_args[0][0]
        assert runtime_request.config_snapshot == {}
        assert PRECONVERSION_INTERMEDIATES_OPTION not in runtime_request.options

    def test_preconversion_rejects_untrusted_config_before_backend_execution(
        self,
        mock_runtime: MagicMock,
        mock_config: MagicMock,
        tmp_path,
    ) -> None:
        source = tmp_path / "legacy.doc"
        source.write_text("legacy")
        mock_config.snapshot.side_effect = RuntimeError("configuration state is untrusted")
        ctrl = ApplicationController(runtime_port=mock_runtime, config_port=mock_config)
        request = ConversionRequest(
            request_id="untrusted-preconversion",
            input_refs=[_file_ref(str(source), "doc")],
            target_format="md",
        )

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=["docx", "md"],
            ),
            patch("docwen_application.preconversion.pre_converter.pre_convert") as pre_convert_mock,
            pytest.raises(RuntimeError, match="configuration state is untrusted"),
        ):
            ctrl.execute_single(request)

        mock_config.snapshot.assert_called_once_with()
        mock_config.get.assert_not_called()
        pre_convert_mock.assert_not_called()
        mock_runtime.execute.assert_not_called()
