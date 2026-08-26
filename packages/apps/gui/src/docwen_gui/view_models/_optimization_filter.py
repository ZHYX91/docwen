"""GUI projection of the canonical Runtime optimization-resource catalog.

Runtime owns resource existence, bindings, actions and route options.  The
``optimize`` config file is intentionally only a user policy overlay: it may
disable known resources and order them, but it cannot invent capabilities.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any, Literal

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.optimization_catalog import (
    OptimizationBinding,
    OptimizationCatalog,
    OptimizationResource,
    inspect_optimization_catalog,
)

logger = logging.getLogger(__name__)


@dataclass(frozen=True, slots=True)
class OptimizationSource:
    """One canonical input identity used to match an optimization binding."""

    detected_format: str
    source_category: str

    def __post_init__(self) -> None:
        object.__setattr__(self, "detected_format", self.detected_format.strip().lower())
        object.__setattr__(self, "source_category", self.source_category.strip().lower())


@dataclass(frozen=True, slots=True)
class OptimizationChoice:
    """One selectable public resource and its resolved execution facts."""

    id: str
    label: str
    action_name: str
    bindings: tuple[OptimizationBinding, ...]
    route_options: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class OptimizationChoicesResult:
    """Explicit discovery state; a ready empty list is not a failure."""

    status: Literal["ready", "failed"]
    choices: tuple[OptimizationChoice, ...]
    error: CapabilityUnavailableError | None = None

    @property
    def items(self) -> dict[str, str]:
        return {choice.id: choice.label for choice in self.choices}

    def get(self, resource_id: str | None) -> OptimizationChoice | None:
        if not resource_id:
            return None
        return next((choice for choice in self.choices if choice.id == resource_id), None)


def _failed(message: str, error: Exception | None = None) -> OptimizationChoicesResult:
    failure = error if isinstance(error, CapabilityUnavailableError) else CapabilityUnavailableError(message)
    return OptimizationChoicesResult(status="failed", choices=(), error=failure)


def _load_catalog(controller: Any) -> tuple[OptimizationCatalog | None, CapabilityUnavailableError | None]:
    if controller is None:
        return None, CapabilityUnavailableError("Runtime optimization discovery is unavailable.")
    describe = getattr(controller, "describe_runtime_capabilities", None)
    if not callable(describe):
        return None, CapabilityUnavailableError("Runtime optimization discovery is unavailable.")
    try:
        projection = describe()
    except Exception as exc:  # the consumer must contain discovery failures
        failure = exc if isinstance(exc, CapabilityUnavailableError) else CapabilityUnavailableError(str(exc))
        return None, failure
    if not isinstance(projection, dict):
        return None, CapabilityUnavailableError("Runtime optimization discovery returned an invalid projection.")
    inspected = inspect_optimization_catalog(projection)
    return inspected.catalog, inspected.error


def _optimization_policy(controller: Any) -> tuple[tuple[str, ...], frozenset[str]]:
    """Return configured order and disabled IDs without treating config as a registry."""

    cfg_port = getattr(controller, "config_port", None)
    if cfg_port is None:
        return (), frozenset()
    try:
        raw = cfg_port.get("optimize", {})
    except Exception:
        logger.warning(
            "Config read failed; using empty policy (key=optimize, stage=optimization-policy)", exc_info=True
        )
        return (), frozenset()
    if not isinstance(raw, dict):
        return (), frozenset()
    settings = raw.get("settings")
    raw_order = settings.get("order") if isinstance(settings, dict) else None
    order = tuple(item for item in raw_order if isinstance(item, str)) if isinstance(raw_order, list) else ()
    raw_types = raw.get("types")
    disabled: set[str] = set()
    if isinstance(raw_types, dict):
        for resource_id, policy in raw_types.items():
            if isinstance(resource_id, str) and isinstance(policy, dict) and policy.get("enabled") is False:
                disabled.add(resource_id)
    return order, frozenset(disabled)


def _ordered_resources(
    catalog: OptimizationCatalog, configured_order: tuple[str, ...]
) -> tuple[OptimizationResource, ...]:
    by_id = {resource.id: resource for resource in catalog.resources}
    ordered_ids = [resource_id for resource_id in configured_order if resource_id in by_id]
    ordered_ids.extend(resource.id for resource in catalog.resources if resource.id not in ordered_ids)
    return tuple(by_id[resource_id] for resource_id in ordered_ids)


def _display_name(resource: OptimizationResource, locale: str) -> str:
    del locale  # Locale selects translated text; it never hides a capability.
    try:
        from docwen_gui.i18n import t

        return str(t(f"cli.interactive.optimization_types.{resource.id}", default=resource.name) or resource.name)
    except Exception:
        logger.warning("Optimization label lookup failed; using catalog name (stage=display-name)", exc_info=True)
        return resource.name


def _binding_for_source(
    resource: OptimizationResource,
    source: OptimizationSource,
    *,
    target: str,
) -> OptimizationBinding | None:
    available = tuple(binding for binding in resource.bindings if binding.available and binding.target == target)
    exact = next((binding for binding in available if binding.source == source.detected_format), None)
    if exact is not None:
        return exact
    if not source.detected_format:
        return next((binding for binding in available if binding.source_category == source.source_category), None)
    return next(
        (
            binding
            for binding in available
            if binding.source == source.source_category and binding.source_category == source.source_category
        ),
        None,
    )


def _route_option_intersection(
    catalog: OptimizationCatalog,
    bindings: tuple[OptimizationBinding, ...],
) -> tuple[str, ...]:
    if not bindings:
        return ()
    common = set(catalog.options_for_route(bindings[0].route_id))
    for binding in bindings[1:]:
        common.intersection_update(catalog.options_for_route(binding.route_id))
    return tuple(option for option in catalog.options_for_route(bindings[0].route_id) if option in common)


def discover_optimization_choices(
    controller: Any,
    *,
    locale: str,
    sources: tuple[OptimizationSource, ...] = (),
    target: str = "md",
) -> OptimizationChoicesResult:
    """Project canonical Runtime resources through optional local policy.

    Settings supplies a category-only source; the operation panel supplies
    exact detected formats.  Every selected input must have an available
    binding.  For a batch, route options are the intersection of all matches.
    """

    catalog, error = _load_catalog(controller)
    if catalog is None:
        return _failed("Runtime optimization discovery failed.", error)

    configured_order, disabled = _optimization_policy(controller)
    choices: list[OptimizationChoice] = []
    for resource in _ordered_resources(catalog, configured_order):
        if resource.id in disabled or not resource.available:
            continue
        if sources:
            matched: list[OptimizationBinding] = []
            for source in sources:
                binding = _binding_for_source(resource, source, target=target)
                if binding is None:
                    break
                matched.append(binding)
            else:
                bindings = tuple(matched)
            if len(matched) != len(sources):
                continue
        else:
            bindings = tuple(binding for binding in resource.bindings if binding.available and binding.target == target)
            if not bindings:
                continue
        choices.append(
            OptimizationChoice(
                id=resource.id,
                label=_display_name(resource, locale),
                action_name=resource.action_name,
                bindings=bindings,
                route_options=_route_option_intersection(catalog, bindings),
            )
        )
    return OptimizationChoicesResult(status="ready", choices=tuple(choices))


__all__ = [
    "OptimizationChoice",
    "OptimizationChoicesResult",
    "OptimizationSource",
    "discover_optimization_choices",
]
