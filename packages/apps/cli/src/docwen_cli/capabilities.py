"""Runtime-owned capability discovery helpers for the protocol 3 CLI."""

from __future__ import annotations

from typing import Any

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.optimization_catalog import OptimizationCatalog, parse_optimization_catalog
from docwen_application.runtime_capability_catalog import (
    RuntimeCapabilityCatalog,
    parse_runtime_capability_catalog,
)


def normalize_target_format(target_format: str | None) -> str:
    """Return the exact public target identifier without normalization."""
    return str(target_format or "")


def runtime_capability_projection(controller: Any | None) -> dict[str, Any]:
    """Read and validate the canonical runtime capability projection.

    Discovery commands must consume the composition root's projection instead
    of rebuilding source/gate tables inside each command.
    """

    if controller is None:
        raise CapabilityUnavailableError("Runtime capability discovery requires an initialized runtime.")
    describe = getattr(controller, "describe_runtime_capabilities", None)
    if not callable(describe):
        raise CapabilityUnavailableError("Runtime capability discovery is unavailable in this CLI assembly.")
    try:
        projection = describe()
    except CapabilityUnavailableError:
        raise
    except Exception as exc:
        raise CapabilityUnavailableError("Runtime capability discovery failed.") from exc
    if not isinstance(projection, dict):
        raise CapabilityUnavailableError("Runtime capability discovery returned an invalid projection.")
    contract = projection.get("contract")
    runtime = projection.get("runtime")
    security = projection.get("security")
    gates = projection.get("gates")
    sources = projection.get("sources")
    counts = projection.get("counts")
    if (
        contract != {"id": "docwen.runtime-capabilities", "version": 1}
        or not isinstance(runtime, dict)
        or runtime.get("state") != "available"
        or not isinstance(security, dict)
        or not isinstance(security.get("dependency_egress_guard"), dict)
        or not isinstance(gates, list)
        or not isinstance(sources, list)
        or not isinstance(counts, dict)
    ):
        raise CapabilityUnavailableError("Runtime capability discovery returned an incomplete projection.")
    return projection


def runtime_optimization_catalog(controller: Any | None) -> OptimizationCatalog:
    """Return the shared typed optimization catalog for this runtime."""

    return parse_optimization_catalog(runtime_capability_projection(controller))


def runtime_route_catalog(controller: Any | None) -> RuntimeCapabilityCatalog:
    """Return the shared typed route catalog for this runtime."""

    return parse_runtime_capability_catalog(runtime_capability_projection(controller))


def runtime_optimization_projection(controller: Any | None) -> dict[str, object]:
    """Return the canonical optimization projection after shared validation."""

    return runtime_optimization_catalog(controller).to_dict()


def resolve_optimization_action(controller: Any | None, resource_id: str) -> str:
    """Resolve a public optimization ID to its declared internal action."""

    match = runtime_optimization_catalog(controller).get(resource_id)
    if match is None:
        raise ValueError(f"Unknown optimization resource: {resource_id}")
    return match.action_name
