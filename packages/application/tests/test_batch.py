"""Unit tests for AggregateCommand and AggregateWorkflow.

Tests verify:
- AggregateCommand validates aggregate actions
- AggregateWorkflow passes all input_refs in one port.execute() call
- AggregateWorkflow correctly validates >=2 input_refs
- AggregateWorkflow rejects non-aggregate actions
- summary() returns correct aggregate-shaped counts
"""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from docwen_application.ports.runtime import RuntimePort
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest

pytestmark = pytest.mark.unit


# ── Fixtures ────────────────────────────────────────────────────────────────


@pytest.fixture
def mock_runtime() -> MagicMock:
    return MagicMock(spec=RuntimePort)


# ── Helpers ─────────────────────────────────────────────────────────────────


def _file_ref(path: str, fmt: str = "pdf") -> FileRef:
    return FileRef(path=path, format=fmt, category="layout")


def _make_aggregate_request(
    request_id: str,
    *paths: str,
    action_name: str = "merge_pdfs",
    target_format: str = "pdf",
) -> ConversionRequest:
    return ConversionRequest(
        request_id=request_id,
        input_refs=[_file_ref(p) for p in paths],
        target_format=target_format,
        action_name=action_name,
    )


# ── AggregateCommand ────────────────────────────────────────────────────────


class TestAggregateCommand:
    """AggregateCommand critical path tests."""

    def test_construction_with_valid_action(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import AggregateCommand

        cmd = AggregateCommand(mock_runtime, action_name="merge_pdfs")
        assert cmd.action_name == "merge_pdfs"

    def test_construction_rejects_unknown_action(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import AggregateCommand

        with pytest.raises(ValueError, match="not a known aggregate action"):
            AggregateCommand(mock_runtime, action_name="convert")

    def test_construction_rejects_empty_action(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import AggregateCommand

        with pytest.raises(ValueError, match="not a known aggregate action"):
            AggregateCommand(mock_runtime, action_name="")

    def test_all_known_actions_accepted(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import (
            AGGREGATE_ACTIONS,
            AggregateCommand,
        )

        for action in AGGREGATE_ACTIONS:
            cmd = AggregateCommand(mock_runtime, action_name=action)
            assert cmd.action_name == action

    def test_execute_delegates_to_runtime(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import AggregateCommand
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="a1", success=True)
        mock_runtime.execute.return_value = expected

        cmd = AggregateCommand(mock_runtime, action_name="merge_pdfs")
        request = _make_aggregate_request("a1", "/a.pdf", "/b.pdf")
        result = cmd.execute(request)

        mock_runtime.execute.assert_called_once_with(request)
        assert result is expected

    def test_execute_passes_all_input_refs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import AggregateCommand
        from docwen_core.models.result import ConversionResult

        mock_runtime.execute.return_value = ConversionResult(task_id="a2", success=True)

        cmd = AggregateCommand(mock_runtime, action_name="merge_tables")
        request = _make_aggregate_request(
            "a2",
            "/t1.xlsx",
            "/t2.xlsx",
            "/t3.xlsx",
            action_name="merge_tables",
            target_format="xlsx",
        )
        cmd.execute(request)

        # Verify the runtime received the request with ALL input_refs intact
        called_request = mock_runtime.execute.call_args[0][0]
        assert len(called_request.input_refs) == 3

    def test_rejects_single_input(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import AggregateCommand

        cmd = AggregateCommand(mock_runtime, action_name="merge_images_to_tiff")
        request = _make_aggregate_request(
            "a3",
            "/img.png",
            action_name="merge_images_to_tiff",
        )

        with pytest.raises(ValueError, match="at least two"):
            cmd.execute(request)

    def test_failure_result_propagated(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import AggregateCommand
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        failure = ConversionResult(
            task_id="a4",
            success=False,
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message="merge error",
            ),
        )
        mock_runtime.execute.return_value = failure

        cmd = AggregateCommand(mock_runtime, action_name="merge_pdfs")
        request = _make_aggregate_request("a4", "/a.pdf", "/b.pdf")
        result = cmd.execute(request)

        assert not result.success
        assert result.error.error_type == "conversion_failed"


# ── AggregateWorkflow ────────────────────────────────────────────────────────


class TestAggregateWorkflow:
    """AggregateWorkflow critical path tests."""

    def test_execute_with_two_inputs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="aw1", success=True)
        mock_runtime.execute.return_value = expected

        wf = AggregateWorkflow(mock_runtime, action_name="merge_pdfs")
        request = _make_aggregate_request("aw1", "/a.pdf", "/b.pdf")
        result = wf.execute(request)

        mock_runtime.execute.assert_called_once_with(request)
        assert result is expected

    def test_execute_with_three_inputs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="aw2", success=True)
        mock_runtime.execute.return_value = expected

        wf = AggregateWorkflow(mock_runtime, action_name="merge_tables")
        request = _make_aggregate_request(
            "aw2",
            "/t1.xlsx",
            "/t2.xlsx",
            "/t3.xlsx",
            action_name="merge_tables",
            target_format="xlsx",
        )
        result = wf.execute(request)

        mock_runtime.execute.assert_called_once()
        assert result is expected

    def test_rejects_single_input(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow

        wf = AggregateWorkflow(mock_runtime, action_name="merge_pdfs")
        request = _make_aggregate_request("aw3", "/only.pdf")

        with pytest.raises(ValueError, match="at least two"):
            wf.execute(request)

    def test_rejects_zero_inputs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow

        wf = AggregateWorkflow(mock_runtime, action_name="merge_pdfs")
        request = _make_aggregate_request("aw4")  # no paths

        with pytest.raises(ValueError, match="at least two"):
            wf.execute(request)

    def test_rejects_non_aggregate_action(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow

        wf = AggregateWorkflow(mock_runtime, action_name="merge_pdfs")
        request = _make_aggregate_request(
            "aw5",
            "/a.pdf",
            "/b.pdf",
            action_name="convert",  # not an aggregate action
        )

        with pytest.raises(ValueError, match="requires an aggregate action"):
            wf.execute(request)

    def test_falls_back_to_constructor_action_name(self, mock_runtime: MagicMock) -> None:
        """When request.action_name is empty, use the constructor action."""
        from docwen_application.workflows.batch import AggregateWorkflow
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="aw6", success=True)
        mock_runtime.execute.return_value = expected

        wf = AggregateWorkflow(mock_runtime, action_name="merge_images_to_tiff")
        # Request with empty action_name — should fall back to constructor's
        request = _make_aggregate_request(
            "aw6",
            "/a.png",
            "/b.png",
            action_name="",  # empty
        )
        result = wf.execute(request)

        assert result.success
        mock_runtime.execute.assert_called_once()

    def test_all_aggregate_actions_work(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow
        from docwen_core.models.result import ConversionResult

        mock_runtime.execute.return_value = ConversionResult(task_id="aw7", success=True)

        for action in ["merge_pdfs", "merge_tables", "merge_images_to_tiff"]:
            wf = AggregateWorkflow(mock_runtime, action_name=action)
            request = _make_aggregate_request(
                "aw7",
                "/a.pdf",
                "/b.pdf",
                action_name=action,
            )
            result = wf.execute(request)
            assert result.success

    def test_runtime_not_called_on_validation_error(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow

        wf = AggregateWorkflow(mock_runtime, action_name="merge_pdfs")
        request = _make_aggregate_request("aw8", "/only.pdf")

        with pytest.raises(ValueError):
            wf.execute(request)
        mock_runtime.execute.assert_not_called()

    # ── summary() ────────────────────────────────────────────────────

    def test_summary_success(self) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow
        from docwen_core.models.result import ConversionResult

        wf = AggregateWorkflow(MagicMock(), action_name="merge_pdfs")
        s = wf.summary(ConversionResult(task_id="s1", success=True))
        assert s == {"total": 1, "success": 1, "failed": 0, "skipped": 0, "cancelled": 0}

    def test_summary_failure(self) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow
        from docwen_core.models.result import ConversionResult

        wf = AggregateWorkflow(MagicMock(), action_name="merge_pdfs")
        s = wf.summary(ConversionResult(task_id="s2", success=False))
        assert s == {"total": 1, "success": 0, "failed": 1, "skipped": 0, "cancelled": 0}

    # ── Properties ────────────────────────────────────────────────────

    def test_action_name_property(self) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow

        wf = AggregateWorkflow(MagicMock(), action_name="merge_tables")
        assert wf.action_name == "merge_tables"

    def test_events_property(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import AggregateWorkflow

        wf = AggregateWorkflow(mock_runtime, action_name="merge_pdfs")
        assert wf.events == []
        assert isinstance(wf.events, list)


# ── is_aggregate_action helper ──────────────────────────────────────────────


class TestIsAggregateAction:
    """Tests for the ``is_aggregate_action`` sentinel function."""

    def test_known_actions(self) -> None:
        from docwen_application.commands.batch import is_aggregate_action

        assert is_aggregate_action("merge_pdfs") is True
        assert is_aggregate_action("merge_tables") is True
        assert is_aggregate_action("merge_images_to_tiff") is True

    def test_unknown_actions(self) -> None:
        from docwen_application.commands.batch import is_aggregate_action

        assert is_aggregate_action("convert") is False
        assert is_aggregate_action("validate") is False
        assert is_aggregate_action("") is False
        assert is_aggregate_action("merge_pdf") is False  # no trailing s

    def test_aggregate_actions_constant(self) -> None:
        from docwen_application.commands.batch import AGGREGATE_ACTIONS

        assert len(AGGREGATE_ACTIONS) == 3
        assert "merge_pdfs" in AGGREGATE_ACTIONS
        assert "merge_tables" in AGGREGATE_ACTIONS
        assert "merge_images_to_tiff" in AGGREGATE_ACTIONS
