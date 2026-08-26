"""ConvertCommand — application-layer command for single-file conversion."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_application.ports.runtime import RuntimePort


class ConvertCommand:
    """Command to execute a single-file conversion.

    This is the application-layer entry point that GUI and CLI both use.
    It selects the appropriate workflow and delegates to the runtime port.
    """

    def __init__(self, runtime_port: RuntimePort) -> None:
        self._runtime = runtime_port

    def execute(self, request: Any) -> Any:
        """Execute a single-file conversion.

        Args:
            request: A ``ConversionRequest`` with exactly one source and any
                typed resources declared by the route.

        Returns:
            A ``ConversionResult``.
        """
        from docwen_application.workflows.single_file import SingleFileWorkflow

        workflow = SingleFileWorkflow(self._runtime)
        return workflow.execute(request)
