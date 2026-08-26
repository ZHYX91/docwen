"""Unit tests for application workflows — SingleFileWorkflow and BatchWorkflow."""

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


def _file_ref(path: str, fmt: str = "txt") -> FileRef:
    return FileRef(path=path, format=fmt, category="document")


def _make_request(
    request_id: str,
    *paths: str,
    target_format: str = "md",
    action_name: str = "",
    options: dict | None = None,
) -> ConversionRequest:
    return ConversionRequest(
        request_id=request_id,
        input_refs=[_file_ref(p) for p in paths],
        target_format=target_format,
        action_name=action_name,
        options=options or {},
    )


# ── SingleFileWorkflow ──────────────────────────────────────────────────────


class TestSingleFileWorkflow:
    """SingleFileWorkflow critical path tests."""

    def test_execute_delegates_to_runtime(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.single_file import SingleFileWorkflow
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="r1", success=True)
        mock_runtime.execute.return_value = expected

        wf = SingleFileWorkflow(mock_runtime)
        request = _make_request("r1", "/f.txt")
        result = wf.execute(request)

        mock_runtime.execute.assert_called_once_with(request)
        assert result is expected

    def test_rejects_empty_input_refs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.single_file import SingleFileWorkflow

        wf = SingleFileWorkflow(mock_runtime)
        request = _make_request("r1")  # no input_refs

        with pytest.raises(ValueError, match="at least one"):
            wf.execute(request)

    def test_rejects_multiple_input_refs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.single_file import SingleFileWorkflow

        wf = SingleFileWorkflow(mock_runtime)
        request = _make_request("r1", "/a.txt", "/b.txt")

        with pytest.raises(ValueError, match="BatchWorkflow"):
            wf.execute(request)

    def test_accepts_one_source_with_declared_resource(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.single_file import SingleFileWorkflow
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="r1", success=True)
        mock_runtime.execute.return_value = expected
        request = _make_request("r1", "/source.md", "/image.png")
        request.input_refs[0].input_kind = "document"
        request.input_refs[0].input_role = "source"
        request.input_refs[1].input_kind = "resource"
        request.input_refs[1].input_role = "linked_resource"

        result = SingleFileWorkflow(mock_runtime).execute(request)

        mock_runtime.execute.assert_called_once_with(request)
        assert result is expected

    def test_accepts_neutral_document_with_numbering_plan(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.single_file import SingleFileWorkflow
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="r1", success=True)
        mock_runtime.execute.return_value = expected
        request = _make_request("r1", "/resolved-document.json", "/numbering-export-plan.json")
        request.input_refs[0].input_kind = "document"
        request.input_refs[0].input_role = "neutral_document"
        request.input_refs[1].input_kind = "resource"
        request.input_refs[1].input_role = "numbering_export_plan"

        result = SingleFileWorkflow(mock_runtime).execute(request)

        mock_runtime.execute.assert_called_once_with(request)
        assert result is expected

    def test_rejects_numbering_plan_without_primary_document(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.single_file import SingleFileWorkflow

        request = _make_request("r1", "/numbering-export-plan.json")
        request.input_refs[0].input_kind = "resource"
        request.input_refs[0].input_role = "numbering_export_plan"

        with pytest.raises(ValueError, match="exactly one primary document input"):
            SingleFileWorkflow(mock_runtime).execute(request)

        mock_runtime.execute.assert_not_called()

    def test_events_starts_empty(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.single_file import SingleFileWorkflow

        wf = SingleFileWorkflow(mock_runtime)
        assert wf.events == []

    def test_runtime_not_called_on_validation_error(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.single_file import SingleFileWorkflow

        wf = SingleFileWorkflow(mock_runtime)
        request = _make_request("r1")

        with pytest.raises(ValueError):
            wf.execute(request)
        mock_runtime.execute.assert_not_called()


# ── BatchWorkflow ───────────────────────────────────────────────────────────


class TestBatchWorkflow:
    """BatchWorkflow critical path tests."""

    def test_execute_with_multiple_inputs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import BatchWorkflow
        from docwen_core.models.result import ConversionResult

        mock_runtime.execute.side_effect = [
            ConversionResult(task_id="b1-0", success=True),
            ConversionResult(task_id="b1-1", success=True),
            ConversionResult(task_id="b1-2", success=True),
        ]

        wf = BatchWorkflow(mock_runtime)
        request = _make_request("b1", "/f0.txt", "/f1.txt", "/f2.txt")
        results = wf.execute(request)

        assert len(results) == 3
        assert mock_runtime.execute.call_count == 3

    def test_continue_on_error_false_stops_on_failure(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import BatchWorkflow
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        ok = ConversionResult(task_id="b2-0", success=True)
        failure = ConversionResult(
            task_id="b2-1",
            success=False,
            error=ConversionErrorInfo(error_type="conversion_failed", message="fail"),
        )
        mock_runtime.execute.side_effect = [ok, failure]

        wf = BatchWorkflow(mock_runtime, continue_on_error=False)
        request = _make_request("b2", "/f0.txt", "/f1.txt", "/f2.txt")
        results = wf.execute(request)

        assert len(results) == 3
        assert results[0].success
        assert not results[1].success
        assert results[2].error.error_type == "skipped"
        # Runtime only called twice (third was skipped)
        assert mock_runtime.execute.call_count == 2

    def test_continue_on_error_true_continues(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import BatchWorkflow
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        mock_runtime.execute.side_effect = [
            ConversionResult(
                task_id="b3-0",
                success=False,
                error=ConversionErrorInfo(error_type="conversion_failed", message="fail1"),
            ),
            ConversionResult(task_id="b3-1", success=True),
            ConversionResult(
                task_id="b3-2",
                success=False,
                error=ConversionErrorInfo(error_type="conversion_failed", message="fail2"),
            ),
        ]

        wf = BatchWorkflow(mock_runtime, continue_on_error=True)
        request = _make_request("b3", "/f0.txt", "/f1.txt", "/f2.txt")
        results = wf.execute(request)

        assert len(results) == 3
        assert not results[0].success
        assert results[1].success
        assert not results[2].success
        assert mock_runtime.execute.call_count == 3

    def test_rejects_empty_input_refs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import BatchWorkflow

        wf = BatchWorkflow(mock_runtime)
        request = _make_request("b4")

        with pytest.raises(ValueError, match="at least one"):
            wf.execute(request)

    def test_single_input_is_valid(self, mock_runtime: MagicMock) -> None:
        """BatchWorkflow should accept a single file for batching consistency."""
        from docwen_application.workflows.batch import BatchWorkflow
        from docwen_core.models.result import ConversionResult

        mock_runtime.execute.return_value = ConversionResult(task_id="b5-0", success=True)

        wf = BatchWorkflow(mock_runtime)
        request = _make_request("b5", "/single.txt")
        results = wf.execute(request)

        assert len(results) == 1
        assert results[0].success

    # ── summary() ──────────────────────────────────────────────────────

    def test_summary_all_success(self) -> None:
        from docwen_application.workflows.batch import BatchWorkflow
        from docwen_core.models.result import ConversionResult

        wf = BatchWorkflow(MagicMock())
        results = [ConversionResult(task_id=f"t{i}", success=True) for i in range(5)]
        s = wf.summary(results)

        assert s == {"total": 5, "success": 5, "failed": 0, "skipped": 0, "cancelled": 0}

    def test_summary_mixed(self) -> None:
        from docwen_application.workflows.batch import BatchWorkflow
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        wf = BatchWorkflow(MagicMock())
        results = [
            ConversionResult(task_id="t0", success=True),
            ConversionResult(
                task_id="t1",
                success=False,
                error=ConversionErrorInfo(error_type="conversion_failed", message="fail"),
            ),
            ConversionResult(
                task_id="t2",
                success=False,
                error=ConversionErrorInfo(error_type="skipped", message="skip"),
            ),
            ConversionResult(
                task_id="t3",
                success=False,
                error=ConversionErrorInfo(error_type="cancelled", message="cancel"),
            ),
        ]
        s = wf.summary(results)

        assert s == {"total": 4, "success": 1, "failed": 1, "skipped": 1, "cancelled": 1}

    def test_summary_empty(self) -> None:
        from docwen_application.workflows.batch import BatchWorkflow

        wf = BatchWorkflow(MagicMock())
        s = wf.summary([])
        assert s == {"total": 0, "success": 0, "failed": 0, "skipped": 0, "cancelled": 0}

    # ── Properties ─────────────────────────────────────────────────────

    def test_events_property(self, mock_runtime: MagicMock) -> None:
        from docwen_application.workflows.batch import BatchWorkflow

        wf = BatchWorkflow(mock_runtime)
        assert wf.events == []
        assert isinstance(wf.events, list)

    def test_continue_on_error_property(self) -> None:
        from docwen_application.workflows.batch import BatchWorkflow

        wf1 = BatchWorkflow(MagicMock())
        assert wf1.continue_on_error is True

        wf2 = BatchWorkflow(MagicMock(), continue_on_error=False)
        assert wf2.continue_on_error is False
