"""Focused tests split from test_runtime_config_wiring.py."""

from __future__ import annotations

import pytest

from ._runtime_config_wiring_support import (
    PROJECT_CONFIGS,
    Any,
    Path,
    _dummy_result,
    build_document_style_catalog,
    dataclass,
    logging,
    tempfile,
)

pytestmark = pytest.mark.unit


class TestConfigFlowToExecutionContext:
    """Verify config flows end-to-end through the runtime pipeline."""

    def test_adapter_populates_config_snapshot_from_loader(self, tmp_path: Path) -> None:
        """Adapter with config_loader must populate config_snapshot on request."""
        from docwen_runtime.adapters import RuntimePortAdapter

        # Build a minimal TaskManager that records what it receives
        class _Recorder:
            def __init__(self) -> None:
                self.requests: list[Any] = []

            def execute_single(self, req: Any, on_event: Any = None) -> Any:
                self.requests.append(req)
                return _dummy_result(req.request_id)

            def execute_batch(self, req: Any, on_event: Any = None) -> list[Any]:
                self.requests.append(req)
                return [_dummy_result(req.request_id)]

            def cancel(self, task_id: str) -> None:
                pass

            def cancel_all(self) -> int:
                return 0

        from docwen_runtime.config.loader import ConfigLoader

        with tempfile.TemporaryDirectory() as tmpdir:
            config_loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            recorder = _Recorder()
            adapter = RuntimePortAdapter(
                recorder,  # type: ignore[arg-type]
                config_loader=config_loader,
            )

            from docwen_core.models.file_ref import FileRef
            from docwen_core.models.request import ConversionRequest, OutputPolicy

            input_path = tmp_path / "config-input.md"
            input_path.write_text("# Config input\n", encoding="utf-8")
            req = ConversionRequest(
                request_id="test-001",
                input_refs=[FileRef(path=str(input_path), format="markdown", category="markdown")],
                target_format="document",
                action_name="validate",
                options={},
                output_policy=OutputPolicy(),
                # config_snapshot is empty — adapter should populate it
            )

            adapter.execute(req)

            # Request received by TaskManager must have config_snapshot populated
            recorded = recorder.requests[0]
            assert recorded.config_snapshot, "Adapter must populate config_snapshot from ConfigLoader"
            assert "gui" in recorded.config_snapshot
            assert "link" in recorded.config_snapshot
            assert "proofread" in recorded.config_snapshot
            assert recorded.config_snapshot["link"]["non_embed_links"]["wiki_mode"] == "hyperlink"

    def test_adapter_without_config_loader_leaves_snapshot_empty(self, tmp_path: Path) -> None:
        """Without config_loader, the adapter should not populate snapshot."""
        from docwen_runtime.adapters import RuntimePortAdapter

        class _Recorder:
            def __init__(self) -> None:
                self.requests: list[Any] = []

            def execute_single(self, req: Any, on_event: Any = None) -> Any:
                self.requests.append(req)
                return _dummy_result(req.request_id)

            def execute_batch(self, req: Any, on_event: Any = None) -> list[Any]:
                self.requests.append(req)
                return [_dummy_result(req.request_id)]

            def cancel(self, task_id: str) -> None:
                pass

            def cancel_all(self) -> int:
                return 0

        recorder = _Recorder()
        adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]
        # No config_loader

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        input_path = tmp_path / "no-loader-input.md"
        input_path.write_text("# No loader\n", encoding="utf-8")
        req = ConversionRequest(
            request_id="test-002",
            input_refs=[FileRef(path=str(input_path), format="markdown", category="markdown")],
            target_format="document",
            action_name="validate",
            options={},
            output_policy=OutputPolicy(),
        )

        adapter.execute(req)
        recorded = recorder.requests[0]
        assert not recorded.config_snapshot, "Without config_loader, snapshot should remain empty"

    def test_adapter_sends_multi_input_aggregate_to_single_execution(self, tmp_path: Path) -> None:
        """Aggregate actions must keep all input refs together for merge-capable plugins."""
        from docwen_runtime.adapters import RuntimePortAdapter

        class _Recorder:
            def __init__(self) -> None:
                self.single_requests: list[Any] = []
                self.batch_requests: list[Any] = []

            def execute_single(self, req: Any, on_event: Any = None) -> Any:
                self.single_requests.append(req)
                return _dummy_result(req.request_id)

            def execute_batch(self, req: Any, on_event: Any = None) -> list[Any]:
                self.batch_requests.append(req)
                return [_dummy_result(req.request_id)]

            def cancel(self, task_id: str) -> None:
                pass

            def cancel_all(self) -> int:
                return 0

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        recorder = _Recorder()
        adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]
        first = tmp_path / "a.pdf"
        second = tmp_path / "b.pdf"
        first.write_bytes(b"%PDF-1.4\nfirst\n")
        second.write_bytes(b"%PDF-1.4\nsecond\n")
        req = ConversionRequest(
            request_id="merge-001",
            input_refs=[
                FileRef(path=str(first), format="pdf", category="layout"),
                FileRef(path=str(second), format="pdf", category="layout"),
            ],
            target_format="pdf",
            action_name="merge_pdfs",
            options={},
            output_policy=OutputPolicy(),
        )

        adapter.execute(req)

        assert recorder.batch_requests == []
        assert [request.request_id for request in recorder.single_requests] == [req.request_id]
        assert len(recorder.single_requests[0].input_refs) == 2

    def test_adapter_sends_multi_input_non_aggregate_to_batch_execution(self, tmp_path: Path) -> None:
        """Regular multi-file conversion should still use batch execution."""
        from docwen_runtime.adapters import RuntimePortAdapter

        class _Recorder:
            def __init__(self) -> None:
                self.single_requests: list[Any] = []
                self.batch_requests: list[Any] = []

            def execute_single(self, req: Any, on_event: Any = None) -> Any:
                self.single_requests.append(req)
                return _dummy_result(req.request_id)

            def execute_batch(self, req: Any, on_event: Any = None) -> list[Any]:
                self.batch_requests.append(req)
                return [_dummy_result(req.request_id)]

            def cancel(self, task_id: str) -> None:
                pass

            def cancel_all(self) -> int:
                return 0

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy

        recorder = _Recorder()
        adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]
        first = tmp_path / "a.md"
        second = tmp_path / "b.md"
        first.write_text("# First\n", encoding="utf-8")
        second.write_text("# Second\n", encoding="utf-8")
        req = ConversionRequest(
            request_id="batch-001",
            input_refs=[
                FileRef(path=str(first), format="markdown", category="markdown"),
                FileRef(path=str(second), format="markdown", category="markdown"),
            ],
            target_format="docx",
            action_name="",
            options={},
            output_policy=OutputPolicy(),
        )

        adapter.execute(req)

        assert recorder.single_requests == []
        assert [request.request_id for request in recorder.batch_requests] == [req.request_id]

    def test_adapter_preserves_existing_config_snapshot(self, tmp_path: Path) -> None:
        """If request already has config_snapshot, adapter must NOT overwrite."""
        from docwen_runtime.adapters import RuntimePortAdapter

        class _Recorder:
            def __init__(self) -> None:
                self.requests: list[Any] = []

            def execute_single(self, req: Any, on_event: Any = None) -> Any:
                self.requests.append(req)
                return _dummy_result(req.request_id)

            def execute_batch(self, req: Any, on_event: Any = None) -> list[Any]:
                self.requests.append(req)
                return [_dummy_result(req.request_id)]

            def cancel(self, task_id: str) -> None:
                pass

            def cancel_all(self) -> int:
                return 0

        from docwen_runtime.config.loader import ConfigLoader

        with tempfile.TemporaryDirectory() as tmpdir:
            config_loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            recorder = _Recorder()
            adapter = RuntimePortAdapter(
                recorder,  # type: ignore[arg-type]
                config_loader=config_loader,
            )

            from docwen_core.models.file_ref import FileRef
            from docwen_core.models.request import ConversionRequest, OutputPolicy

            input_path = tmp_path / "existing-snapshot.md"
            input_path.write_text("# Existing snapshot\n", encoding="utf-8")
            custom_snapshot = {"custom_key": "custom_value"}
            req = ConversionRequest(
                request_id="test-003",
                input_refs=[FileRef(path=str(input_path), format="markdown", category="markdown")],
                target_format="document",
                action_name="validate",
                options={},
                output_policy=OutputPolicy(),
                config_snapshot=custom_snapshot,
            )

            adapter.execute(req)
            recorded = recorder.requests[0]
            assert recorded.config_snapshot == custom_snapshot, "Existing config_snapshot must not be overwritten"

    def test_config_flow_to_execution_context(self) -> None:
        """Verify that config_snapshot reaches RuntimeExecutionContext."""
        from docwen_core.cancellation import CancellationToken
        from docwen_core.export_semantics import MarkdownExportSemantics
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_runtime._execution_context import RuntimeExecutionContext

        config_snapshot = {
            "engine": {"enable_symbol_pairing": False},
            "gui": {"window": {"default_mode": "batch"}},
        }

        req = ConversionRequest(
            request_id="test-flow",
            input_refs=[FileRef(path="/tmp/test.docx", format="docx", category="document")],
            target_format="document",
            action_name="validate",
            options={},
            output_policy=OutputPolicy(),
            config_snapshot=config_snapshot,
        )

        @dataclass
        class _FakeWorkspace:
            input_path: str = "/tmp/test.docx"
            staging_dir: str = "/tmp/staging"

            def create_artifact_path(self, kind: str, suffix: str) -> str:
                return f"/tmp/staging/{kind}_1{suffix}"

            def add_artifact(self, manifest: Any) -> None:
                pass

        ctx = RuntimeExecutionContext(
            request=req,
            workspace=_FakeWorkspace(),  # type: ignore[arg-type]
            config_snapshot=config_snapshot,
            cancellation_token=CancellationToken(),
            document_style_catalog=build_document_style_catalog(
                config_snapshot,
                locales_dir=PROJECT_CONFIGS.parent / "i18n" / "locales",
            ),
            markdown_export_semantics=MarkdownExportSemantics.from_config_snapshot(config_snapshot),
        )

        # Config must be readable through the context
        # ReadOnlyConfigView.get() returns dicts for nested keys
        gui = ctx.config.get("gui", {})
        assert isinstance(gui, dict)
        window = gui.get("window", {})
        assert isinstance(window, dict)
        assert window.get("default_mode") == "batch"
        assert ctx.config.get("gui.window.default_mode") == "batch"
        engine = ctx.config.get("engine", {})
        assert isinstance(engine, dict)
        assert engine.get("enable_symbol_pairing") is False
        assert ctx.config.get("engine.enable_symbol_pairing") is False

    def test_runtime_context_exposes_injected_proofread_rules(self) -> None:
        """RuntimeExecutionContext must expose the pure proofread rule bundle."""
        from docwen_core.cancellation import CancellationToken
        from docwen_core.export_semantics import MarkdownExportSemantics
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.proofread import ProofreadRules
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_runtime._execution_context import RuntimeExecutionContext

        @dataclass
        class _FakeWorkspace:
            input_path: str = "/tmp/test.docx"
            staging_dir: str = "/tmp/staging"

            def create_artifact_path(self, kind: str, suffix: str) -> str:
                return f"/tmp/staging/{kind}_1{suffix}"

            def add_artifact(self, manifest: Any) -> None:
                pass

        rules = ProofreadRules(typos_map={"已": ("己",)})
        req = ConversionRequest(
            request_id="test-proofread-rules",
            input_refs=[FileRef(path="/tmp/test.docx", format="docx", category="document")],
            target_format="document",
            action_name="validate",
            options={},
            output_policy=OutputPolicy(),
        )

        ctx = RuntimeExecutionContext(
            request=req,
            workspace=_FakeWorkspace(),  # type: ignore[arg-type]
            config_snapshot={},
            cancellation_token=CancellationToken(),
            proofread_rules=rules,
            document_style_catalog=build_document_style_catalog(
                {},
                locales_dir=PROJECT_CONFIGS.parent / "i18n" / "locales",
            ),
            markdown_export_semantics=MarkdownExportSemantics(),
        )

        assert ctx.proofread_rules is rules

    def test_build_proofread_rules_uses_defaults_and_snapshot_values(self) -> None:
        """build_proofread_rules must normalize snapshot data."""
        from docwen_runtime.config import build_proofread_rules

        rules = build_proofread_rules(
            {
                "proofread": {
                    "pairs": {"items": [["<", ">"]]},
                    "symbol_map": {"entries": {"0": ["０"]}},
                    "typos": {"entries": {"已": ["己"]}},
                    "sensitive_words": {"entries": {"机密": ["公开稿"]}},
                }
            }
        )

        assert rules.symbol_pairs == (("<", ">"),)
        assert rules.typos_map == {"已": ("己",)}
        assert rules.sensitive_words == {"机密": ("公开稿",)}
        assert rules.symbol_map["0"] == ("０",)


class TestLoggingPreInit:
    """Verify pre_init_logging() sets up console-only logging."""

    def test_pre_init_returns_docwen_logger(self) -> None:
        from docwen_runtime.logging import pre_init_logging

        logger = pre_init_logging("INFO")
        assert logger.name == "docwen"
        assert logger.level == logging.INFO

    def test_pre_init_adds_console_handler(self) -> None:
        from docwen_runtime.logging import pre_init_logging

        logger = pre_init_logging("DEBUG")
        console_handlers = [h for h in logger.handlers if isinstance(h, logging.StreamHandler)]
        assert len(console_handlers) >= 1

    def test_pre_init_is_idempotent(self) -> None:
        """Calling pre_init_logging() twice should not duplicate handlers."""
        from docwen_runtime.logging import pre_init_logging

        logger = pre_init_logging("INFO")
        handler_count = len(logger.handlers)

        logger2 = pre_init_logging("DEBUG")
        # Should have same handler count (old ones removed)
        assert len(logger2.handlers) == handler_count
