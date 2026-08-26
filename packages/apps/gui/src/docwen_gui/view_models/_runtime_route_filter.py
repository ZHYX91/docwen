"""Typed GUI consumer of the canonical Runtime route catalog."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.runtime_capability_catalog import (
    RouteOperation,
    RuntimeCapabilityCatalog,
    RuntimeRoute,
    parse_runtime_capability_catalog,
)


@dataclass(frozen=True, slots=True)
class RuntimeRouteSource:
    """One content-first source identity used by Runtime route resolution."""

    detected_format: str
    source_category: str

    def __post_init__(self) -> None:
        object.__setattr__(self, "detected_format", self.detected_format.strip().lower())
        object.__setattr__(self, "source_category", self.source_category.strip().lower())


@dataclass(frozen=True, slots=True)
class RuntimeCatalogResult:
    """A loaded catalog, a valid catalog with no available routes, or failure."""

    status: Literal["ready", "empty", "failed"]
    catalog: RuntimeCapabilityCatalog | None
    projection: dict[str, object] | None
    error: CapabilityUnavailableError | None = None


@dataclass(frozen=True, slots=True)
class RuntimeRouteChoice:
    """One target common to every selected source and its canonical options."""

    target: str
    routes: tuple[RuntimeRoute, ...]
    options: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class RuntimeRouteChoicesResult:
    """Explicit ready/empty/failed projection for one UI operation surface."""

    status: Literal["ready", "empty", "failed"]
    choices: tuple[RuntimeRouteChoice, ...]
    error: CapabilityUnavailableError | None = None

    @property
    def targets(self) -> tuple[str, ...]:
        return tuple(choice.target for choice in self.choices)

    def get(self, target: str) -> RuntimeRouteChoice | None:
        normalized = str(target or "").strip().lower()
        return next((choice for choice in self.choices if choice.target == normalized), None)

    def select_targets(self, targets: set[str] | frozenset[str]) -> RuntimeRouteChoicesResult:
        if self.status == "failed":
            return self
        normalized = {target.strip().lower() for target in targets if target.strip()}
        choices = tuple(choice for choice in self.choices if choice.target in normalized)
        return RuntimeRouteChoicesResult(status="ready" if choices else "empty", choices=choices)


def load_runtime_catalog(controller: Any) -> RuntimeCatalogResult:
    """Load and validate Runtime's projection without reconstructing route facts."""

    describe = getattr(controller, "describe_runtime_capabilities", None)
    if not callable(describe):
        error = CapabilityUnavailableError("Runtime capability discovery is unavailable.")
        return RuntimeCatalogResult(status="failed", catalog=None, projection=None, error=error)
    try:
        projection = describe()
        if not isinstance(projection, dict):
            raise CapabilityUnavailableError("Runtime capability discovery returned an invalid projection.")
        catalog = parse_runtime_capability_catalog(projection)
    except Exception as exc:
        error = exc if isinstance(exc, CapabilityUnavailableError) else CapabilityUnavailableError(str(exc))
        return RuntimeCatalogResult(status="failed", catalog=None, projection=None, error=error)
    has_available_route = any(route.available for source in catalog.sources for route in source.routes)
    return RuntimeCatalogResult(
        status="ready" if has_available_route else "empty",
        catalog=catalog,
        projection=projection,
    )


def _common_options(routes: tuple[RuntimeRoute, ...]) -> tuple[str, ...]:
    if not routes:
        return ()
    common = set(routes[0].options)
    for route in routes[1:]:
        common.intersection_update(route.options)
    return tuple(option for option in routes[0].options if option in common)


def discover_runtime_route_choices(
    controller: Any,
    *,
    sources: tuple[RuntimeRouteSource, ...],
    operation: RouteOperation,
    action_name: str = "",
) -> RuntimeRouteChoicesResult:
    """Return available targets common to all sources using Runtime precedence."""

    loaded = load_runtime_catalog(controller)
    if loaded.status == "failed" or loaded.catalog is None:
        return RuntimeRouteChoicesResult(status="failed", choices=(), error=loaded.error)
    if not sources:
        return RuntimeRouteChoicesResult(status="empty", choices=())

    normalized_action = action_name.strip().lower()
    per_source: list[dict[str, RuntimeRoute]] = []
    for source in sources:
        for source_id in loaded.catalog.source_ids_for(source.detected_format, source.source_category):
            source_group = loaded.catalog.source(source_id)
            if source_group is None:
                continue
            signatures: set[tuple[RouteOperation, str, str]] = set()
            for candidate in source_group.routes:
                if candidate.signature in signatures:
                    error = CapabilityUnavailableError(
                        f"Runtime capability discovery returned duplicate route signatures in source {source_id}."
                    )
                    return RuntimeRouteChoicesResult(status="failed", choices=(), error=error)
                signatures.add(candidate.signature)
        routes: dict[str, RuntimeRoute] = {}
        for route in loaded.catalog.effective_routes(source.detected_format, source.source_category):
            if not route.available or route.operation != operation or route.action_name != normalized_action:
                continue
            if route.target in routes:
                error = CapabilityUnavailableError(
                    "Runtime capability discovery returned ambiguous routes for "
                    f"{source.detected_format or source.source_category} -> {route.target} "
                    f"(action={normalized_action!r})."
                )
                return RuntimeRouteChoicesResult(status="failed", choices=(), error=error)
            routes[route.target] = route
        per_source.append(routes)

    first_routes = per_source[0]
    common_targets = set(first_routes)
    for routes in per_source[1:]:
        common_targets.intersection_update(routes)
    choices = tuple(
        RuntimeRouteChoice(
            target=target,
            routes=tuple(routes[target] for routes in per_source),
            options=_common_options(tuple(routes[target] for routes in per_source)),
        )
        for target in first_routes
        if target in common_targets
    )
    return RuntimeRouteChoicesResult(status="ready" if choices else "empty", choices=choices)


__all__ = [
    "RuntimeCatalogResult",
    "RuntimeRouteChoice",
    "RuntimeRouteChoicesResult",
    "RuntimeRouteSource",
    "discover_runtime_route_choices",
    "load_runtime_catalog",
]
