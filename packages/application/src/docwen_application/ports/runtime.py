"""Runtime port — the interface application requires from the runtime layer.

Defined HERE in application (Dependency Inversion Principle):
application defines WHAT it needs; runtime/bundle provides the IMPLEMENTATION.
"""

from typing import Any, Protocol, runtime_checkable


@runtime_checkable
class RuntimePort(Protocol):
    """The interface through which application delegates execution to the runtime.

    The protocol intentionally remains structural so CLI, GUI, tests, and
    bundle composition can inject runtime implementations without importing
    runtime internals into the application layer.
    """

    def execute(self, request: Any) -> Any:
        """Execute a conversion request and return the result.

        Args:
            request: A ConversionRequest (defined in docwen_core.models).

        Returns:
            A ConversionResult.
        """
        ...

    def cancel(self, task_id: str) -> None:
        """Request cancellation of a running task.

        Args:
            task_id: The identifier of the task to cancel.
        """
        ...

    def shutdown(self) -> None:
        """Release runtime-owned resources and reject further lifecycle use."""
        ...

    @property
    def is_available(self) -> bool:
        """Whether the runtime is initialized and ready to accept requests."""
        ...


@runtime_checkable
class CapabilityDiscoveryPort(Protocol):
    """Optional read-only reflection surface implemented by production runtime."""

    def describe_capabilities(self) -> dict[str, Any]:
        """Return the current loaded composition and machine gate projection."""
        ...


@runtime_checkable
class CancellationReservationPort(Protocol):
    """Optional Runtime capability for one Application-owned task lifetime.

    A reservation makes cancellation atomic across the short gaps before
    Runtime token registration and after synchronous execution returns.
    Structural RuntimePort implementations that do not expose this optional
    capability retain the base port contract.
    """

    def reserve_cancellation(self, task_id: str) -> None:
        """Reserve Runtime cancellation state before task admission."""
        ...

    def release_cancellation(self, task_id: str) -> None:
        """Release reserved cancellation state after Application handoff."""
        ...


@runtime_checkable
class OutputManifestPersistencePort(Protocol):
    """Optional best-effort sidecar capability implemented by production Runtime."""

    def persist_output_manifests(self, request: Any, result: Any) -> Any:
        """Return terminal result(s) with configured manifest artifacts attached."""
        ...


@runtime_checkable
class ArtifactBundleCommitPort(Protocol):
    """Runtime-owned integrity commit for one semantic BundleDraft."""

    def commit(self, *, task_id: str, staging_root: str, draft: Any) -> Any:
        """Validate paths and graph, hash deliverables, and return Artifact Bundle v2."""
        ...

    def discard(self, *, staging_root: str, artifact_paths: list[str]) -> None:
        """Remove only the named rejected deliverables from request-owned staging."""
        ...


@runtime_checkable
class ConfigPort(Protocol):
    """The interface through which application reads configuration.

    Application does NOT care where config comes from (TOML files,
    env vars, CLI args, or in-memory defaults) — only that it can
    read typed values.
    """

    def get(self, key: str, default: Any = None) -> Any:
        """Read a configuration value by dotted key path.

        Args:
            key: Dotted path, e.g. "output.mode".
            default: Value returned when the key is not found.

        Returns:
            The configuration value, or default.

        Raises:
            RuntimeError: Effective state is untrusted because the last reconciliation
                did not complete.
        """
        ...

    def snapshot(self) -> dict[str, Any]:
        """Copy effective state or raise when the last reconciliation did not complete."""
        ...

    def set(self, key: str, value: Any) -> bool:
        """Persist one value all-or-nothing when handled-failure compensation succeeds."""
        ...

    def set_many(self, values: dict[str, Any]) -> bool:
        """Persist a batch all-or-nothing for handled in-process failures.

        When compensation succeeds, a ``False`` result leaves every affected
        user file at its pre-operation image.  Compensation or reconciliation
        failures remain a hard failure.  If the reconciliation reload itself
        does not complete, ``get`` and ``snapshot`` may raise instead of
        exposing stale state.  This contract does not imply cross-process
        coordination or crash/power-loss atomicity.
        """
        ...

    def get_file_text(self, rel_path: str) -> str | None:
        """Read one registered editable config source as effective TOML text."""
        ...

    def save_file_text(self, rel_path: str, content: str) -> bool:
        """Persist TOML all-or-nothing when handled-failure compensation succeeds."""
        ...

    def reset_file(self, rel_path: str) -> bool:
        """Reset one file all-or-nothing when handled-failure compensation succeeds."""
        ...

    def reset_group(self, group: str) -> bool:
        """Reset a group all-or-nothing when handled-failure compensation succeeds."""
        ...

    def reset_all(self) -> bool:
        """Reset all files all-or-nothing when handled-failure compensation succeeds."""
        ...

    def reload(self) -> None: ...


@runtime_checkable
class PresenterPort(Protocol):
    """The interface through which application sends results for display.

    Application does NOT know whether the result will be rendered as
    text (CLI), widgets (GUI), or JSON (API). The presenter port
    abstracts the output medium.
    """

    def present_result(self, result: Any) -> None:
        """Present a successful conversion result to the user.

        Args:
            result: A ConversionResult (defined in docwen_core.models).
        """
        ...

    def present_error(self, task_id: str, error: Exception) -> None:
        """Present a conversion error to the user.

        Args:
            task_id: The task that failed.
            error: The exception that caused the failure.
        """
        ...
