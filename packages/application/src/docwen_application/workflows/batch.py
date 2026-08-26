"""BatchWorkflow and AggregateWorkflow — orchestrates batch and merge operations.

BatchWorkflow: per-file batch conversion (many-to-many).
    Each file is converted independently via individual calls to
    ``runtime_port.execute()``.

AggregateWorkflow: merge/aggregate operations (many-to-one).
    All input files are passed together in a single ``port.execute()`` call
    so the runtime can feed them to the appropriate merge converter.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_application.ports.runtime import RuntimePort


class BatchWorkflow:
    """Workflow for converting multiple files as a batch.

    Each file in the batch is converted independently via individual
    calls to ``runtime_port.execute()``.  The workflow **owns** the
    batch semantics: it iterates over input files, builds single-file
    requests, and decides when to skip remaining items.

    Responsibilities:
    - Validate the batch request.
    - Split into single-file sub-requests and call the runtime port.
    - Aggregate results (success/failure/skip counts).
    - Relay task events to the presenter.

    This workflow does NOT:
    - Resolve plugins or routes directly.
    - Manage workspace or output paths.
    - Update GUI widgets or CLI output directly.
    """

    def __init__(
        self,
        runtime_port: RuntimePort,
        *,
        continue_on_error: bool = True,
    ) -> None:
        self._runtime = runtime_port
        self._continue_on_error = continue_on_error
        self._events: list[Any] = []

    def execute(self, request: Any) -> list[Any]:
        """Execute a batch conversion.

        Iterates over ``request.input_refs`` and calls
        ``runtime_port.execute()`` once per file.  Each call is
        independent — a failure in one file does not affect others
        unless *continue_on_error* is ``False``.

        Args:
            request: A ``ConversionRequest`` with one or more ``input_refs``.

        Returns:
            A ``list[ConversionResult]``, one per input file.

        Raises:
            ValueError: If the request has zero input refs.
        """
        if not hasattr(request, "input_refs") or len(request.input_refs) == 0:
            raise ValueError("ConversionRequest must have at least one input file")

        # Build single-file sub-requests
        from docwen_core.models.request import ConversionRequest

        results: list[Any] = []
        for i, input_ref in enumerate(request.input_refs):
            single_request = ConversionRequest(
                request_id=f"{request.request_id}-{i}",
                input_refs=[input_ref],
                target_format=request.target_format,
                action_name=getattr(request, "action_name", ""),
                options=dict(getattr(request, "options", {})),
                output_policy=getattr(request, "output_policy", None),  # pyright: ignore[reportArgumentType]
                config_snapshot=dict(getattr(request, "config_snapshot", {})),
            )
            result = self._runtime.execute(single_request)
            results.append(result)

            if not getattr(result, "success", False) and not self._continue_on_error:
                # Mark remaining as skipped
                self._append_skipped(results, request.request_id, request.input_refs[i + 1 :])
                break

        return results

    def _append_skipped(self, results: list[Any], request_id: str, remaining: list[Any]) -> None:
        """Append skipped results for remaining input files."""
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        for _ in remaining:
            results.append(
                ConversionResult(
                    task_id=f"{request_id}-skipped",
                    success=False,
                    error=ConversionErrorInfo(
                        error_type="skipped",
                        message="Skipped due to previous error",
                    ),
                )
            )

    @property
    def events(self) -> list[Any]:
        """Return collected task events from the last execution."""
        return list(self._events)

    def summary(self, results: list[Any]) -> dict[str, int]:
        """Compute a summary of batch results.

        Args:
            results: List of ``ConversionResult``.

        Returns:
            Dict with keys ``total``, ``success``, ``failed``, ``skipped``, ``cancelled``.
        """
        total = len(results)
        success = sum(1 for r in results if getattr(r, "success", False))
        failed = sum(
            1
            for r in results
            if not getattr(r, "success", False)
            and getattr(getattr(r, "error", None), "error_type", "") not in ("skipped", "cancelled")
        )
        skipped = sum(1 for r in results if getattr(getattr(r, "error", None), "error_type", "") == "skipped")
        cancelled = sum(1 for r in results if getattr(getattr(r, "error", None), "error_type", "") == "cancelled")
        return {
            "total": total,
            "success": success,
            "failed": failed,
            "skipped": skipped,
            "cancelled": cancelled,
        }

    @property
    def continue_on_error(self) -> bool:
        return self._continue_on_error


# ── Aggregate workflow ─────────────────────────────────────────────────


class AggregateWorkflow:
    """Workflow for aggregate (merge) operations — many inputs → one output.

    Unlike ``BatchWorkflow``, this workflow does NOT split the request
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

        from docwen_application.commands.batch import is_aggregate_action

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
