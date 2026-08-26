"""Unit tests for application commands — ConvertCommand and BatchCommand."""

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


def _file_ref(path: str, fmt: str = "txt") -> FileRef:
    return FileRef(path=path, format=fmt, category="document")


def _request(request_id: str, *paths: str, target_format: str = "md") -> ConversionRequest:
    return ConversionRequest(
        request_id=request_id,
        input_refs=[_file_ref(p) for p in paths],
        target_format=target_format,
    )


# ── ConvertCommand ──────────────────────────────────────────────────────────


class TestConvertCommand:
    """ConvertCommand minimal critical path tests."""

    def test_construction_stores_runtime(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.convert import ConvertCommand

        cmd = ConvertCommand(mock_runtime)
        assert cmd._runtime is mock_runtime

    def test_execute_returns_result(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.convert import ConvertCommand
        from docwen_core.models.result import ConversionResult

        expected = ConversionResult(task_id="t1", success=True)
        mock_runtime.execute.return_value = expected

        cmd = ConvertCommand(mock_runtime)
        request = _request("t1", "/input.txt")
        result = cmd.execute(request)

        mock_runtime.execute.assert_called_once_with(request)
        assert result is expected

    def test_execute_validates_single_input(self, mock_runtime: MagicMock) -> None:
        """SingleFileWorkflow rejects zero or multiple input_refs."""
        from docwen_application.commands.convert import ConvertCommand

        cmd = ConvertCommand(mock_runtime)

        # Zero inputs
        request = _request("t1")  # no paths = empty input_refs
        with pytest.raises(ValueError, match="at least one"):
            cmd.execute(request)

        # Multiple inputs
        request = _request("t2", "/a.txt", "/b.txt")
        with pytest.raises(ValueError, match="BatchWorkflow"):
            cmd.execute(request)

    def test_failure_result_propagated(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.convert import ConvertCommand
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        failure = ConversionResult(
            task_id="t1",
            success=False,
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message="bad input",
            ),
        )
        mock_runtime.execute.return_value = failure

        cmd = ConvertCommand(mock_runtime)
        request = _request("t1", "/bad.txt")
        result = cmd.execute(request)

        assert not result.success
        assert result.error.error_type == "conversion_failed"


# ── BatchCommand ────────────────────────────────────────────────────────────


class TestBatchCommand:
    """BatchCommand minimal critical path tests."""

    def test_construction_defaults(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import BatchCommand

        cmd = BatchCommand(mock_runtime)
        assert cmd._runtime is mock_runtime
        assert cmd._continue_on_error is True

    def test_construction_continue_on_error_false(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import BatchCommand

        cmd = BatchCommand(mock_runtime, continue_on_error=False)
        assert not cmd._continue_on_error

    def test_execute_batch_with_multiple_inputs(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import BatchCommand
        from docwen_core.models.result import ConversionResult

        r0 = ConversionResult(task_id="b1-0", success=True)
        r1 = ConversionResult(task_id="b1-1", success=True)
        r2 = ConversionResult(task_id="b1-2", success=True)
        mock_runtime.execute.side_effect = [r0, r1, r2]

        cmd = BatchCommand(mock_runtime)
        request = _request("b1", "/f0.txt", "/f1.txt", "/f2.txt")
        results = cmd.execute(request)

        assert len(results) == 3
        assert all(r.success for r in results)
        assert mock_runtime.execute.call_count == 3

    def test_execute_continue_on_error_stops(self, mock_runtime: MagicMock) -> None:
        """continue_on_error=False halts on first failure, marking rest skipped."""
        from docwen_application.commands.batch import BatchCommand
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        failure = ConversionResult(
            task_id="b2-1",
            success=False,
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message="fail",
            ),
        )
        mock_runtime.execute.return_value = failure

        cmd = BatchCommand(mock_runtime, continue_on_error=False)
        request = _request("b2", "/f0.txt", "/f1.txt", "/f2.txt")
        results = cmd.execute(request)

        # 3 results: 1 failure + 2 skipped
        assert len(results) == 3
        assert not results[0].success
        assert results[0].error.error_type == "conversion_failed"
        assert results[1].error.error_type == "skipped"
        assert results[2].error.error_type == "skipped"

    def test_execute_continue_on_error_true_continues(self, mock_runtime: MagicMock) -> None:
        """continue_on_error=True does NOT skip after a failure."""
        from docwen_application.commands.batch import BatchCommand
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        failure = ConversionResult(
            task_id="b3-0",
            success=False,
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message="fail",
            ),
        )
        ok = ConversionResult(task_id="b3-1", success=True)
        mock_runtime.execute.side_effect = [failure, ok]

        cmd = BatchCommand(mock_runtime, continue_on_error=True)
        request = _request("b3", "/f0.txt", "/f1.txt")
        results = cmd.execute(request)

        assert len(results) == 2
        assert not results[0].success
        assert results[1].success

    def test_execute_validates_has_input(self, mock_runtime: MagicMock) -> None:
        from docwen_application.commands.batch import BatchCommand

        cmd = BatchCommand(mock_runtime)
        request = _request("b4")  # no paths = empty input_refs
        with pytest.raises(ValueError, match="at least one"):
            cmd.execute(request)
