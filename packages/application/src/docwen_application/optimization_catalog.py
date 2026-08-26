"""Typed consumer view of the runtime optimization-resource contract."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from types import MappingProxyType
from typing import Literal, NoReturn, cast

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.runtime_capability_catalog import (
    RuntimeCapabilityCatalog,
    parse_runtime_capability_catalog,
)

OptimizationState = Literal["available", "available_with_limits", "unavailable"]
CatalogStatus = Literal["ready", "failed"]

_RESOURCE_FIELDS = {"id", "name", "action_name", "scopes", "available", "state", "bindings"}
_BINDING_FIELDS = {
    "scope",
    "route_id",
    "source",
    "source_category",
    "target",
    "available",
    "state",
}
_VALID_STATES = {"available", "available_with_limits", "unavailable"}


def _fail(message: str) -> NoReturn:
    raise CapabilityUnavailableError(message)


def _text(value: object, *, field: str) -> str:
    if not isinstance(value, str) or not value:
        _fail(f"Runtime optimization discovery returned an invalid {field}.")
    return value


def _state(value: object, *, field: str) -> OptimizationState:
    if value not in _VALID_STATES:
        _fail(f"Runtime optimization discovery returned an invalid {field}.")
    return cast(OptimizationState, value)


@dataclass(frozen=True, slots=True)
class OptimizationBinding:
    scope: str
    route_id: str
    source: str
    source_category: str
    target: str
    available: bool
    state: OptimizationState

    def to_dict(self) -> dict[str, object]:
        return {
            "scope": self.scope,
            "route_id": self.route_id,
            "source": self.source,
            "source_category": self.source_category,
            "target": self.target,
            "available": self.available,
            "state": self.state,
        }


@dataclass(frozen=True, slots=True)
class OptimizationResource:
    id: str
    name: str
    action_name: str
    scopes: tuple[str, ...]
    available: bool
    state: OptimizationState
    bindings: tuple[OptimizationBinding, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "id": self.id,
            "name": self.name,
            "action_name": self.action_name,
            "scopes": list(self.scopes),
            "available": self.available,
            "state": self.state,
            "bindings": [binding.to_dict() for binding in self.bindings],
        }


@dataclass(frozen=True, slots=True)
class OptimizationCounts:
    resources: int
    available_resources: int
    unavailable_resources: int
    bindings: int
    available_bindings: int
    unavailable_bindings: int

    def to_dict(self) -> dict[str, int]:
        return {
            "resources": self.resources,
            "available_resources": self.available_resources,
            "unavailable_resources": self.unavailable_resources,
            "bindings": self.bindings,
            "available_bindings": self.available_bindings,
            "unavailable_bindings": self.unavailable_bindings,
        }


@dataclass(frozen=True, slots=True)
class OptimizationCatalog:
    contract_id: str
    contract_version: int
    runtime_state: str
    platform: str
    resources: tuple[OptimizationResource, ...]
    counts: OptimizationCounts
    route_options_by_id: Mapping[str, tuple[str, ...]]

    def get(self, resource_id: str) -> OptimizationResource | None:
        return next((resource for resource in self.resources if resource.id == resource_id), None)

    def compatible_bindings(
        self,
        resource_id: str,
        *,
        source: str,
        target: str,
        available_only: bool = True,
    ) -> tuple[OptimizationBinding, ...]:
        resource = self.get(resource_id)
        if resource is None:
            return ()
        return tuple(
            binding
            for binding in resource.bindings
            if binding.source == source and binding.target == target and (binding.available or not available_only)
        )

    def options_for_route(self, route_id: str) -> tuple[str, ...]:
        return self.route_options_by_id.get(route_id, ())

    def to_dict(self) -> dict[str, object]:
        return {
            "resource": "optimizations",
            "contract": {"id": self.contract_id, "version": self.contract_version},
            "runtime": {"state": self.runtime_state, "platform": self.platform},
            "resources": [resource.to_dict() for resource in self.resources],
            "counts": self.counts.to_dict(),
        }


@dataclass(frozen=True, slots=True)
class OptimizationCatalogResult:
    status: CatalogStatus
    catalog: OptimizationCatalog | None
    error: CapabilityUnavailableError | None


def _parse_binding(
    raw: object,
    *,
    scopes: tuple[str, ...],
    action_name: str,
    route_catalog: RuntimeCapabilityCatalog,
) -> tuple[OptimizationBinding, tuple[str, ...]]:
    if not isinstance(raw, dict) or set(raw) != _BINDING_FIELDS:
        _fail("Runtime optimization discovery returned an invalid binding.")
    available = raw.get("available")
    if not isinstance(available, bool):
        _fail("Runtime optimization discovery returned an invalid binding.")
    binding = OptimizationBinding(
        scope=_text(raw.get("scope"), field="binding scope"),
        route_id=_text(raw.get("route_id"), field="binding route id"),
        source=_text(raw.get("source"), field="binding source"),
        source_category=_text(raw.get("source_category"), field="binding source category"),
        target=_text(raw.get("target"), field="binding target"),
        available=available,
        state=_state(raw.get("state"), field="binding state"),
    )
    if binding.scope not in scopes:
        _fail("Runtime optimization discovery returned an invalid binding.")
    route = route_catalog.routes_by_id.get(binding.route_id)
    if route is None:
        _fail(f"Runtime optimization binding references an unknown route: {binding.route_id}")
    if (
        route.operation != "action"
        or route.action_name != action_name
        or route.source != binding.source
        or route.source_category != binding.source_category
        or route.target != binding.target
        or route.available != binding.available
        or route.state != binding.state
    ):
        _fail(f"Runtime optimization binding disagrees with its route: {binding.route_id}")
    return binding, route.options


def _expected_resource_state(bindings: tuple[OptimizationBinding, ...]) -> OptimizationState:
    available_count = sum(binding.available for binding in bindings)
    if available_count == 0:
        return "unavailable"
    if available_count != len(bindings) or any(binding.state != "available" for binding in bindings):
        return "available_with_limits"
    return "available"


def parse_optimization_catalog(capability_projection: Mapping[str, object]) -> OptimizationCatalog:
    """Parse and cross-check ``docwen.optimizations`` against canonical routes.

    A valid initialized composition may contain zero resources. Missing,
    malformed, contradictory, or route-detached data raises the typed
    ``CapabilityUnavailableError`` instead of becoming an empty catalog.
    """

    route_catalog = parse_runtime_capability_catalog(capability_projection)
    base_runtime = capability_projection.get("runtime")
    projection = capability_projection.get("optimizations")
    if not isinstance(projection, dict) or set(projection) != {
        "resource",
        "contract",
        "runtime",
        "resources",
        "counts",
    }:
        _fail("Runtime optimization discovery is unavailable.")
    if (
        projection.get("resource") != "optimizations"
        or projection.get("contract") != {"id": "docwen.optimizations", "version": 1}
        or projection.get("runtime") != base_runtime
    ):
        _fail("Runtime optimization discovery returned an incomplete projection.")

    raw_resources = projection.get("resources")
    raw_counts = projection.get("counts")
    if not isinstance(raw_resources, list) or not isinstance(raw_counts, dict):
        _fail("Runtime optimization discovery returned an incomplete projection.")

    seen_resource_ids: set[str] = set()
    seen_route_ids: set[str] = set()
    route_options: dict[str, tuple[str, ...]] = {}
    resources: list[OptimizationResource] = []
    for raw_resource in raw_resources:
        if not isinstance(raw_resource, dict) or set(raw_resource) != _RESOURCE_FIELDS:
            _fail("Runtime optimization discovery returned an invalid resource.")
        resource_id = _text(raw_resource.get("id"), field="resource id")
        if resource_id in seen_resource_ids:
            _fail(f"Runtime optimization discovery returned a duplicate resource id: {resource_id}")
        seen_resource_ids.add(resource_id)
        action_name = _text(raw_resource.get("action_name"), field="resource action")
        raw_scopes = raw_resource.get("scopes")
        raw_bindings = raw_resource.get("bindings")
        if (
            not isinstance(raw_scopes, list)
            or not raw_scopes
            or not all(isinstance(scope, str) and scope for scope in raw_scopes)
            or len(raw_scopes) != len(set(raw_scopes))
            or not isinstance(raw_bindings, list)
            or not raw_bindings
            or not isinstance(raw_resource.get("available"), bool)
        ):
            _fail("Runtime optimization discovery returned an invalid resource.")
        scopes = tuple(raw_scopes)
        parsed_bindings: list[OptimizationBinding] = []
        for raw_binding in raw_bindings:
            binding, options = _parse_binding(
                raw_binding,
                scopes=scopes,
                action_name=action_name,
                route_catalog=route_catalog,
            )
            if binding.route_id in seen_route_ids:
                _fail(f"Runtime optimization route is bound more than once: {binding.route_id}")
            seen_route_ids.add(binding.route_id)
            route_options[binding.route_id] = options
            parsed_bindings.append(binding)
        bindings = tuple(parsed_bindings)
        if set(scopes) != {binding.scope for binding in bindings}:
            _fail(f"Runtime optimization resource has unbound scopes: {resource_id}")
        state = _state(raw_resource.get("state"), field="resource state")
        available = bool(raw_resource["available"])
        if available != any(binding.available for binding in bindings) or state != _expected_resource_state(bindings):
            _fail(f"Runtime optimization resource state is inconsistent: {resource_id}")
        resources.append(
            OptimizationResource(
                id=resource_id,
                name=_text(raw_resource.get("name"), field="resource name"),
                action_name=action_name,
                scopes=scopes,
                available=available,
                state=state,
                bindings=bindings,
            )
        )

    bindings = tuple(binding for resource in resources for binding in resource.bindings)
    counts = OptimizationCounts(
        resources=len(resources),
        available_resources=sum(resource.available for resource in resources),
        unavailable_resources=sum(not resource.available for resource in resources),
        bindings=len(bindings),
        available_bindings=sum(binding.available for binding in bindings),
        unavailable_bindings=sum(not binding.available for binding in bindings),
    )
    if raw_counts != counts.to_dict():
        _fail("Runtime optimization discovery returned inconsistent counts.")
    return OptimizationCatalog(
        contract_id="docwen.optimizations",
        contract_version=1,
        runtime_state="available",
        platform=route_catalog.platform,
        resources=tuple(resources),
        counts=counts,
        route_options_by_id=MappingProxyType(route_options),
    )


def inspect_optimization_catalog(capability_projection: Mapping[str, object]) -> OptimizationCatalogResult:
    """Return an explicit ready/failed result without folding failure into empty."""

    try:
        catalog = parse_optimization_catalog(capability_projection)
    except CapabilityUnavailableError as exc:
        return OptimizationCatalogResult(status="failed", catalog=None, error=exc)
    return OptimizationCatalogResult(status="ready", catalog=catalog, error=None)


__all__ = [
    "OptimizationBinding",
    "OptimizationCatalog",
    "OptimizationCatalogResult",
    "OptimizationCounts",
    "OptimizationResource",
    "inspect_optimization_catalog",
    "parse_optimization_catalog",
]
