"""Unit tests for application commands — ConvertCommand."""

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
        with pytest.raises(ValueError, match="execute_batch"):
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
