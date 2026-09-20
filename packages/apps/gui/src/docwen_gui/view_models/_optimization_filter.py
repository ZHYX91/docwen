"""GUI projection of the canonical Runtime optimization-resource catalog.

Runtime owns resource existence, bindings, actions and route options.  The
Application layer owns pre-conversion composition.  This view therefore never
invents a GUI-only optimizer route: when an admitted legacy document reaches an
optimizer through the canonical Application pre-conversion chain, discovery
reuses the optimizer's real DOCX binding and action.

The ``optimize`` config file remains only a user policy overlay: it may disable
known resources and order them, but it cannot invent capabilities.
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
from docwen_application.optimization_selection import OptimizationSource, select_optimizations

logger = logging.getLogger(__name__)


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


def _display_name(resource: OptimizationResource, locale: str) -> str:
    del locale  # Locale selects translated text; it never hides a capability.
    try:
        from docwen_gui.i18n import t

        return str(t(f"cli.interactive.optimization_types.{resource.id}", default=resource.name) or resource.name)
    except Exception:
        logger.warning("Optimization label lookup failed; using catalog name (stage=display-name)", exc_info=True)
        return resource.name


def discover_optimization_choices(
    controller: Any,
    *,
    locale: str,
    sources: tuple[OptimizationSource, ...] = (),
    target: str = "md",
) -> OptimizationChoicesResult:
    """Project canonical resources through Application composition and user policy.

    Settings supplies a category-only source; the operation panel supplies
    exact detected formats. Every selected input must resolve to one available
    Runtime binding, directly or through the canonical Application
    pre-conversion chain. For a batch, route options are the intersection of
    all resolved bindings.
    """

    catalog, error = _load_catalog(controller)
    if catalog is None:
        return _failed("Runtime optimization discovery failed.", error)

    configured_order, disabled = _optimization_policy(controller)
    selections = select_optimizations(
        catalog,
        sources=sources,
        target=target,
        configured_order=configured_order,
        disabled=disabled,
    )
    choices = tuple(
        OptimizationChoice(
            id=selection.resource.id,
            label=_display_name(selection.resource, locale),
            action_name=selection.resource.action_name,
            bindings=selection.bindings,
            route_options=selection.route_options,
        )
        for selection in selections
    )
    return OptimizationChoicesResult(status="ready", choices=choices)


__all__ = [
    "OptimizationChoice",
    "OptimizationChoicesResult",
    "discover_optimization_choices",
]
