"""Aggregate workflow for many-to-one operations."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_application.ports.runtime import RuntimePort


class AggregateWorkflow:
    """Workflow for aggregate (merge) operations — many inputs → one output.

    This workflow does NOT split the request
    into single-file sub-requests.  It passes ALL ``input_refs`` together
    in a single ``port.execute()`` call so the runtime can route to a
    merge-capable converter (e.g. ``PdfMerger``, ``TableMergerConverter``,
    ``ImageToTiffMerger``).

    Responsibilities:
    - Validate that the request has at least 2 input refs.
    - Validate that the action is a known aggregate operation.
    - Execute once with the full input_refs list.
    - Return a single ``ConversionResult``.
    """

    def __init__(
        self,
        runtime_port: RuntimePort,
        action_name: str,
    ) -> None:
        self._runtime = runtime_port
        self._action_name = action_name
        self._events: list[Any] = []

    @property
    def action_name(self) -> str:
        return self._action_name

    def execute(self, request: Any) -> Any:
        """Execute an aggregate operation.

        Args:
            request: A ``ConversionRequest`` with two or more ``input_refs``
                and an aggregate ``action_name``.

        Returns:
            A ``ConversionResult`` representing the merged output.

        Raises:
            ValueError: If fewer than 2 input_refs or unrecognized action.
        """
        if not hasattr(request, "input_refs") or len(request.input_refs) < 2:
            raise ValueError(
                "Aggregate operations require at least two input files, "
                f"got {len(request.input_refs) if hasattr(request, 'input_refs') else 0}"
            )

        from docwen_application.commands.aggregate import is_aggregate_action

        action = getattr(request, "action_name", self._action_name) or self._action_name
        if not is_aggregate_action(action):
            raise ValueError(f"AggregateWorkflow requires an aggregate action, got {action!r}")

        # Pass the full request (with ALL input_refs) to the runtime.
        result = self._runtime.execute(request)
        return result

    @property
    def events(self) -> list[Any]:
        """Return collected task events from the last execution."""
        return list(self._events)

    def summary(self, result: Any) -> dict[str, int]:
        """Compute a summary from a single aggregate result.

        Args:
            result: A ``ConversionResult``.

        Returns:
            Dict with keys ``total``, ``success``, ``failed``, ``skipped``, ``cancelled``.
        """
        success = 1 if getattr(result, "success", False) else 0
        return {
            "total": 1,
            "success": success,
            "failed": 0 if success else 1,
            "skipped": 0,
            "cancelled": 0,
        }
