"""Focused tests split from test_controller.py."""

from __future__ import annotations

from ._controller_support import (
    Any,
    ApplicationController,
    ConversionRequest,
    FileRef,
    MagicMock,
    Path,
    _file_ref,
    _request,
    patch,
    pytest,
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

    def test_document_preconversion_receives_configured_priority(
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
        mock_runtime.execute.return_value = ConversionResult(task_id="configured-pre", success=True)
        mock_config.snapshot.return_value = {
            "software": {
                "default_priority": {
                    "word_processors": ["msoffice_word", "libreoffice", "wps_writer"],
                }
            }
        }
        ctrl = ApplicationController(runtime_port=mock_runtime, config_port=mock_config)
        request = ConversionRequest(
            request_id="configured-pre",
            input_refs=[_file_ref(str(source), "doc")],
            target_format="md",
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
                    backend="Microsoft Word",
                ),
            ) as pre_convert_mock,
        ):
            ctrl.execute_single(request)

        assert pre_convert_mock.call_args.kwargs["backend_priority"] == [
            "msoffice_word",
            "libreoffice",
            "wps_writer",
        ]
        mock_config.snapshot.assert_called_once_with()
        mock_config.get.assert_not_called()

    def test_batch_rejects_typed_resources(self, mock_runtime: MagicMock) -> None:
        request = ConversionRequest(
            request_id="typed-batch",
            input_refs=[
                FileRef(path="/source.md", format="markdown", category="markdown"),
                FileRef(
                    path="/bibliography.json",
                    format="resource",
                    category="other",
                    input_role="bibliography",
                    media_type="application/vnd.docwen.semantic-bibliography+json",
                ),
            ],
            target_format="docx",
        )

        with pytest.raises(ValueError, match="only independent source"):
            ApplicationController(runtime_port=mock_runtime).execute_batch(request)

    def test_odt_preconversion_prefers_request_snapshot_over_config_port(
        self,
        mock_runtime: MagicMock,
        mock_config: MagicMock,
        tmp_path,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "legacy.odt"
        source.write_text("legacy")
        preconverted = tmp_path / "legacy_pre.docx"
        preconverted.write_text("hub")
        mock_runtime.execute.return_value = ConversionResult(task_id="snapshot-pre", success=True)
        ctrl = ApplicationController(runtime_port=mock_runtime, config_port=mock_config)
        request = ConversionRequest(
            request_id="snapshot-pre",
            input_refs=[_file_ref(str(source), "odt")],
            target_format="md",
            config_snapshot={"software": {"special_conversions": {"odt": ["libreoffice", "msoffice_word"]}}},
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
                    original_source_format="odt",
                    backend="LibreOffice",
                ),
            ) as pre_convert_mock,
        ):
            ctrl.execute_single(request)

        assert pre_convert_mock.call_args.kwargs["backend_priority"] == [
            "libreoffice",
            "msoffice_word",
        ]
        mock_config.get.assert_not_called()

    def test_execute_single_delegates(self, mock_runtime: MagicMock) -> None:
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="r1", success=True)
        mock_runtime.execute.return_value = expected

        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = _request("r1", "/f.txt")
        result = ctrl.execute_single(request)

        mock_runtime.execute.assert_called_once_with(request)
        assert result is expected

    def test_execute_single_returns_preconversion_failure_without_runtime(self, mock_runtime: MagicMock) -> None:
        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = _request("r1", "/f.doc", source_format="doc")

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

        mock_runtime.execute.assert_not_called()
        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "dependency_missing"

    def test_preconversion_failure_uses_optional_manifest_persistence_port(self) -> None:
        from docwen_core.models.result import ConversionResult

        class _ManifestRuntime:
            def __init__(self) -> None:
                self.persisted: list[tuple[ConversionRequest, ConversionResult]] = []

            def execute(self, request: object) -> object:
                raise AssertionError(f"preconversion failure must not reach runtime: {request!r}")

            def persist_output_manifests(
                self, request: ConversionRequest, result: ConversionResult
            ) -> ConversionResult:
                self.persisted.append((request, result))
                return result

            def cancel(self, task_id: str) -> None:
                del task_id

            def shutdown(self) -> None:
                return None

            @property
            def is_available(self) -> bool:
                return True

        class _ManifestConfig:
            def snapshot(self) -> dict[str, Any]:
                return {"output": {"manifest": {"save_to_output": True, "mask_input_path": True}}}

        runtime = _ManifestRuntime()
        ctrl = ApplicationController(runtime_port=runtime, config_port=_ManifestConfig())  # type: ignore[arg-type]
        request = _request("manifest-preconversion-failure", "/private/legacy.doc", source_format="doc")

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                return_value=["docx", "md"],
            ),
            patch("docwen_application.preconversion.pre_converter.pre_convert", return_value=None),
        ):
            result = ctrl.execute_single(request)

        assert result.success is False
        assert len(runtime.persisted) == 1
        manifest_request, persisted_result = runtime.persisted[0]
        assert persisted_result is result
        assert manifest_request.manifest_context is not None
        assert manifest_request.manifest_context.inputs[0].path == "/private/legacy.doc"
        assert manifest_request.manifest_context.preconversion_steps[0].status == "failed"
        assert manifest_request.manifest_context.preconversion_steps[0].target_format == "docx"

    def test_execute_batch_delegates(self, mock_runtime: MagicMock) -> None:
        from docwen_core.models.result import ConversionResult

        r1 = ConversionResult(task_id="r1-0", success=True)
        r2 = ConversionResult(task_id="r1-1", success=True)
        mock_runtime.execute.side_effect = [r1, r2]

        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = _request("r1", "/f1.txt", "/f2.txt")
        results = ctrl.execute_batch(request)

        assert len(results) == 2
        assert results[0] is r1
        assert results[1] is r2

    def test_execute_batch_preserves_preconversion_failures_in_order(self, mock_runtime: MagicMock) -> None:
        from docwen_core.models.result import ConversionResult

        runtime_success = ConversionResult(task_id="r1-runtime", success=True)
        mock_runtime.execute.return_value = runtime_success

        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = _request(
            "r1",
            "/needs-office.doc",
            "/ready.docx",
            "/also-needs-office.doc",
            source_formats=("doc", "docx", "doc"),
        )

        with (
            patch(
                "docwen_application.preconversion.chain_resolver.resolve_chain",
                side_effect=[
                    ["docx", "md"],
                    ["docx", "md"],
                    ["md"],
                    ["docx", "md"],
                ],
            ),
            patch(
                "docwen_application.preconversion.pre_converter.pre_convert",
                side_effect=[None, None],
            ),
        ):
            results = ctrl.execute_batch(request)

        assert len(results) == 3
        assert results[0].success is False
        assert results[0].error.error_type == "dependency_missing"
        assert results[1] is runtime_success
        assert results[2].success is False
        assert results[2].error.error_type == "dependency_missing"
        mock_runtime.execute.assert_called_once()
        runtime_request = mock_runtime.execute.call_args[0][0]
        assert runtime_request.request_id == "r1-1"
        assert [ref.path for ref in runtime_request.input_refs] == ["/ready.docx"]

    @pytest.mark.parametrize(
        ("actual_format", "path"),
        [
            ("csv", "/table.csv"),
            ("tsv", "/table.tsv"),
        ],
    )
    def test_execute_single_delimited_spreadsheet_to_md_does_not_preconvert(
        self,
        mock_runtime: MagicMock,
        actual_format: str,
        path: str,
    ) -> None:
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id=f"{actual_format}-md", success=True)
        mock_runtime.execute.return_value = expected

        ctrl = ApplicationController(runtime_port=mock_runtime)
        request = _request(
            f"{actual_format}-md",
            path,
            target_format="md",
            source_format=actual_format,
        )

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
        assert runtime_request.input_refs[0].path == path

    def test_typed_resource_is_never_resolved_or_preconverted(
        self,
        mock_runtime: MagicMock,
        tmp_path: Path,
    ) -> None:
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "source.md"
        resource = tmp_path / "bibliography.json"
        source.write_text("# source\n", encoding="utf-8")
        resource.write_text('{"schema":"docwen.semantic_bibliography.v1","entries":[]}', encoding="utf-8")
        expected = ConversionResult(task_id="typed-resource", success=True)
        mock_runtime.execute.return_value = expected
        request = ConversionRequest(
            request_id="typed-resource",
            input_refs=[
                FileRef(
                    path=str(source),
                    format="markdown",
                    category="markdown",
                    input_kind="document",
                    input_role="source",
                    media_type="text/markdown",
                ),
                FileRef(
                    path=str(resource),
                    format="resource",
                    category="other",
                    input_kind="resource",
                    input_role="bibliography",
                    media_type="application/vnd.docwen.semantic-bibliography+json",
                ),
            ],
            target_format="docx",
        )

        def resolve_chain(source_format: str, *_args: object, **_kwargs: object) -> list[str]:
            if source_format == "resource":
                raise AssertionError("typed resource reached route resolution")
            return ["docx"]

        with (
            patch("docwen_application.preconversion.chain_resolver.resolve_chain", side_effect=resolve_chain),
            patch("docwen_application.preconversion.pre_converter.pre_convert") as pre_convert_mock,
        ):
            result = ApplicationController(runtime_port=mock_runtime).execute_single(request)

        assert result is expected
        pre_convert_mock.assert_not_called()
        runtime_request = mock_runtime.execute.call_args[0][0]
        assert runtime_request.input_refs[1] is request.input_refs[1]
        assert runtime_request.input_refs[1].media_type == "application/vnd.docwen.semantic-bibliography+json"

    def test_preconverted_source_keeps_typed_resource_in_single_request(
        self,
        mock_runtime: MagicMock,
        tmp_path: Path,
    ) -> None:
        from docwen_application.preconversion.pre_converter import PreConversionResult
        from docwen_core.models.result import ConversionResult

        source = tmp_path / "legacy.doc"
        converted = tmp_path / "legacy.docx"
        resource = tmp_path / "bibliography.json"
        source.write_text("legacy", encoding="utf-8")
        converted.write_text("converted", encoding="utf-8")
        resource.write_text('{"schema":"docwen.semantic_bibliography.v1","entries":[]}', encoding="utf-8")
        expected = ConversionResult(task_id="typed-preconvert", success=True)
        mock_runtime.execute.return_value = expected
        request = ConversionRequest(
            request_id="typed-preconvert",
            input_refs=[
                FileRef(
                    path=str(source),
                    format="doc",
                    category="document",
                    input_kind="document",
                    input_role="source",
                    media_type="application/msword",
                ),
                FileRef(
                    path=str(resource),
                    format="resource",
                    category="other",
                    input_kind="resource",
                    input_role="bibliography",
                    media_type="application/vnd.docwen.semantic-bibliography+json",
                ),
            ],
            target_format="md",
        )

        def resolve_chain(source_format: str, *_args: object, **_kwargs: object) -> list[str]:
            if source_format == "resource":
                raise AssertionError("typed resource reached route resolution")
            return ["docx", "md"]

        with (
            patch("docwen_application.preconversion.chain_resolver.resolve_chain", side_effect=resolve_chain),
            patch(
                "docwen_application.preconversion.pre_converter.pre_convert",
                return_value=PreConversionResult(
                    pre_converted_path=str(converted),
                    original_source_format="doc",
                    backend="Fake Office",
                ),
            ),
        ):
            result = ApplicationController(runtime_port=mock_runtime).execute_single(request)

        assert result is expected
        mock_runtime.execute.assert_called_once()
        runtime_request = mock_runtime.execute.call_args[0][0]
        assert [item.input_role for item in runtime_request.input_refs] == ["source", "bibliography"]
        assert runtime_request.input_refs[0].format == "docx"
        assert runtime_request.input_refs[0].media_type == (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        assert runtime_request.input_refs[1] is request.input_refs[1]

    def test_preconversion_failure_with_typed_resource_is_single_result(
        self,
        mock_runtime: MagicMock,
        tmp_path: Path,
    ) -> None:
        source = tmp_path / "legacy.doc"
        resource = tmp_path / "bibliography.json"
        source.write_text("legacy", encoding="utf-8")
        resource.write_text("{}", encoding="utf-8")
        request = ConversionRequest(
            request_id="typed-preconvert-failure",
            input_refs=[
                FileRef(path=str(source), format="doc", category="document", input_role="source"),
                FileRef(
                    path=str(resource),
                    format="resource",
                    category="other",
                    input_role="bibliography",
                    media_type="application/vnd.docwen.semantic-bibliography+json",
                ),
            ],
            target_format="md",
        )

        def resolve_chain(source_format: str, *_args: object, **_kwargs: object) -> list[str]:
            if source_format == "resource":
                raise AssertionError("typed resource reached route resolution")
            return ["docx", "md"]

        with (
            patch("docwen_application.preconversion.chain_resolver.resolve_chain", side_effect=resolve_chain),
            patch("docwen_application.preconversion.pre_converter.pre_convert", return_value=None),
        ):
            result = ApplicationController(runtime_port=mock_runtime).execute_single(request)

        assert result.success is False
        assert result.task_id == "typed-preconvert-failure"
        assert result.error is not None
        assert result.error.error_type == "dependency_missing"
        mock_runtime.execute.assert_not_called()
