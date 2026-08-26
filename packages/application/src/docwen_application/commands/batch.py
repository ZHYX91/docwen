"""BatchCommand and AggregateCommand — application-layer commands.

BatchCommand handles per-file batch conversion (many-to-many).
AggregateCommand handles merge/aggregate operations (many-to-one).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_application.ports.runtime import RuntimePort

# ── Known aggregate actions (many-input → single-output) ──────────────────

AGGREGATE_ACTIONS: frozenset[str] = frozenset(
    {
        "merge_pdfs",
        "merge_tables",
        "merge_images_to_tiff",
    }
)


def is_aggregate_action(action_name: str) -> bool:
    """Return True if *action_name* is a known aggregate (merge) operation."""
    return action_name in AGGREGATE_ACTIONS


class BatchCommand:
    """Command to execute a batch conversion.

    This is the application-layer entry point that GUI and CLI both use
    for batch operations.  It selects the ``BatchWorkflow`` and delegates
    to the runtime port.
    """

    def __init__(
        self,
        runtime_port: RuntimePort,
        *,
        continue_on_error: bool = True,
    ) -> None:
        self._runtime = runtime_port
        self._continue_on_error = continue_on_error

    def execute(self, request: Any) -> list[Any]:
        """Execute a batch conversion.

        Args:
            request: A ``ConversionRequest`` with one or more ``input_refs``.

        Returns:
            A ``list[ConversionResult]``, one per input file.
        """
        from docwen_application.workflows.batch import BatchWorkflow

        workflow = BatchWorkflow(
            self._runtime,
            continue_on_error=self._continue_on_error,
        )
        return workflow.execute(request)


class AggregateCommand:
    """Command to execute an aggregate (merge) operation.

    Aggregate operations combine multiple input files into a single output
    (e.g. merge PDFs, merge tables, merge images to TIFF).  Unlike batch
    conversion (which processes each file independently), an aggregate
    command passes ALL input refs in a single ``ConversionRequest`` so the
    runtime can feed them together to the appropriate merge converter.

    Responsibilities:
    - Validate that the action is a known aggregate operation.
    - Delegate to ``AggregateWorkflow`` for execution.
    """

    def __init__(
        self,
        runtime_port: RuntimePort,
        action_name: str,
    ) -> None:
        self._runtime = runtime_port
        self._action_name = action_name
        if not is_aggregate_action(action_name):
            raise ValueError(f"action_name {action_name!r} is not a known aggregate action")

    @property
    def action_name(self) -> str:
        return self._action_name

    def execute(self, request: Any) -> Any:
        """Execute an aggregate conversion.

        Args:
            request: A ``ConversionRequest`` with two or more ``input_refs``
                and an ``action_name`` matching one of the known aggregate
                actions.

        Returns:
            A ``ConversionResult`` representing the merged output.

        Raises:
            ValueError: If the request has fewer than 2 input_refs.
        """
        from docwen_application.workflows.batch import AggregateWorkflow

        workflow = AggregateWorkflow(self._runtime, self._action_name)
        return workflow.execute(request)
