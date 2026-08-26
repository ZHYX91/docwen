"""SingleFileWorkflow — orchestrates a single-file conversion through the runtime port."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_application.ports.runtime import RuntimePort


class SingleFileWorkflow:
    """Workflow for converting a single file.

    Responsibilities:
    - Validate the conversion request.
    - Delegate execution to the runtime port.
    - Collect and relay task events.
    - Return the conversion result.

    This workflow does NOT:
    - Resolve plugins or routes directly.
    - Manage workspace or output paths.
    - Update GUI widgets or CLI output directly.
    """

    def __init__(self, runtime_port: RuntimePort) -> None:
        self._runtime = runtime_port
        self._events: list[Any] = []

    def execute(self, request: Any) -> Any:
        """Execute a single-file conversion.

        Args:
            request: A ``ConversionRequest``.

        Returns:
            A ``ConversionResult``.

        Raises:
            ValueError: If the request has zero or more than one primary document input.
        """
        if not hasattr(request, "input_refs") or len(request.input_refs) == 0:
            raise ValueError("ConversionRequest must have at least one input file")
        primary_inputs = [
            item
            for item in request.input_refs
            if getattr(item, "input_role", "source") in {"source", "neutral_document"}
        ]
        if len(primary_inputs) != 1:
            raise ValueError(
                f"SingleFileWorkflow requires exactly one primary document input, got {len(primary_inputs)}. "
                "Use BatchWorkflow for multiple independent source files."
            )

        return self._runtime.execute(request)

    @property
    def events(self) -> list[Any]:
        """Return collected task events from the last execution."""
        return list(self._events)
