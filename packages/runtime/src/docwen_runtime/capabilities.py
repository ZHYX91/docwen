"""Runtime capability projection from the loaded plugin composition.

Plugin manifests declare routes and typed capability rules.  This module is
the only current-machine evaluator: it probes declared gates and projects one
stable source -> target/action matrix for CLI, GUI, and integration consumers.
"""

from __future__ import annotations

import importlib.util
import sys
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docwen_core.formats.categories import ALL_CATEGORIES, FORMAT_CATEGORY
from docwen_runtime.pymupdf_layout_resources import verify_pymupdf_layout_resource_root

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest, RouteSpec

CAPABILITY_CONTRACT_ID = "docwen.runtime-capabilities"
CAPABILITY_CONTRACT_VERSION = 1
OPTIMIZATION_CONTRACT_ID = "docwen.optimizations"
OPTIMIZATION_CONTRACT_VERSION = 1


class OptimizationContractError(ValueError):
    """Raised when manifest optimization declarations cannot be projected safely."""


@dataclass(frozen=True, slots=True)
class _ModuleGate:
    module: str
    label: str


_MODULE_GATES: dict[str, _ModuleGate] = {
    "python.bs4": _ModuleGate("bs4", "Beautiful Soup"),
    "python.docx": _ModuleGate("docx", "python-docx"),
    "python.easyofd": _ModuleGate("easyofd", "easyofd"),
    "python.ebooklib": _ModuleGate("ebooklib", "EbookLib"),
    "python.fitz": _ModuleGate("fitz", "PyMuPDF"),
    "python.markdownify": _ModuleGate("markdownify", "markdownify"),
    "python.mistune": _ModuleGate("mistune", "Mistune"),
    "python.openpyxl": _ModuleGate("openpyxl", "openpyxl"),
    "python.pdf2docx": _ModuleGate("pdf2docx", "pdf2docx"),
    "python.pillow": _ModuleGate("PIL", "Pillow"),
    "python.pillow_heif": _ModuleGate("pillow_heif", "pillow-heif"),
    "python.pptx": _ModuleGate("pptx", "python-pptx"),
    "python.pymupdf4llm": _ModuleGate("pymupdf4llm", "PyMuPDF4LLM"),
    "python.rapidocr": _ModuleGate("rapidocr_onnxruntime", "RapidOCR"),
}

_OFFICE_PROG_IDS: dict[str, tuple[str, ...]] = {
    "external_office.word": ("Word.Application", "kwps.Application"),
    "external_office.spreadsheet": ("Excel.Application", "ket.Application"),
    "external_office.presentation": ("PowerPoint.Application", "kwpp.Application"),
}

_OFFICE_LABELS: dict[str, str] = {
    "external_office.word": "Word-compatible Office backend",
    "external_office.spreadsheet": "Spreadsheet-compatible Office backend",
    "external_office.presentation": "Presentation-compatible Office backend",
}


def current_platform_id() -> str:
    """Return the platform identifier used by manifest contracts."""

    if sys.platform == "win32":
        return "windows"
    if sys.platform == "darwin":
        return "darwin"
    if sys.platform.startswith("linux"):
        return "linux"
    return sys.platform.lower()


def _module_available(module: str) -> bool:
    try:
        return importlib.util.find_spec(module) is not None
    except (ImportError, ModuleNotFoundError, ValueError):
        return False


def _pymupdf_layout_package_roots() -> tuple[Path, ...]:
    """Return filesystem roots for the import package without exposing them publicly."""
    try:
        spec = importlib.util.find_spec("pymupdf.layout")
    except (ImportError, ModuleNotFoundError, ValueError):
        return ()
    if spec is None:
        return ()
    roots = tuple(Path(value) for value in (spec.submodule_search_locations or ()) if value)
    if roots:
        return roots
    if spec.origin:
        return (Path(spec.origin).parent,)
    return ()


def probe_pymupdf_layout_resources() -> dict[str, Any]:
    """Probe the data files required by the PyMuPDF Layout lazy-load path.

    The result is deliberately path-free so a packaged capability response does
    not disclose the application's internal extraction or installation layout.
    """
    if not _module_available("pymupdf4llm"):
        return {
            "available": False,
            "reason": "module_not_available",
            "resource_types": [],
            "resource_count": 0,
        }

    roots = _pymupdf_layout_package_roots()
    if not roots:
        return {
            "available": False,
            "reason": "resource_package_not_available",
            "resource_types": [],
            "resource_count": 0,
        }

    verifications = tuple(verify_pymupdf_layout_resource_root(root / "resources") for root in roots)
    verification = next(
        (candidate for candidate in verifications if candidate.available),
        max(verifications, key=lambda candidate: candidate.resource_count),
    )
    return {
        "available": verification.available,
        "reason": verification.reason,
        "resource_types": list(verification.resource_types),
        "resource_count": verification.resource_count,
    }


def _registered_com_provider(prog_id: str) -> bool:
    if sys.platform != "win32" or not _module_available("win32com.client"):
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{prog_id}\CLSID"):
            return True
    except (ImportError, OSError):
        return False


def _office_gate_status(gate_id: str) -> dict[str, Any]:
    from docwen_core.office_bridge import find_soffice_path

    providers: list[str] = []
    if find_soffice_path() is not None:
        providers.append("libreoffice")
    providers.extend(prog_id for prog_id in _OFFICE_PROG_IDS[gate_id] if _registered_com_provider(prog_id))
    available = bool(providers)
    return {
        "id": gate_id,
        "kind": "external_office",
        "label": _OFFICE_LABELS[gate_id],
        "available": available,
        "reason": None if available else "no_compatible_backend",
        "providers": providers,
    }


def _gate_status(gate_id: str) -> dict[str, Any]:
    module_gate = _MODULE_GATES.get(gate_id)
    if module_gate is not None:
        if gate_id == "python.pymupdf4llm":
            probe = probe_pymupdf_layout_resources()
            return {
                "id": gate_id,
                "kind": "python_module_with_resources",
                "label": module_gate.label,
                "available": bool(probe["available"]),
                "reason": probe["reason"],
                "module": module_gate.module,
                "resource_types": list(probe["resource_types"]),
                "resource_count": int(probe["resource_count"]),
            }
        available = _module_available(module_gate.module)
        return {
            "id": gate_id,
            "kind": "python_module",
            "label": module_gate.label,
            "available": available,
            "reason": None if available else "module_not_available",
            "module": module_gate.module,
        }
    if gate_id in _OFFICE_PROG_IDS:
        return _office_gate_status(gate_id)
    if gate_id == "backend.pdf_to_docx":
        office = _office_gate_status("external_office.word")
        module_available = _module_available("pdf2docx")
        providers = list(office["providers"])
        if module_available:
            providers.append("pdf2docx")
        available = bool(providers)
        return {
            "id": gate_id,
            "kind": "any_of",
            "label": "PDF to DOCX backend",
            "available": available,
            "reason": None if available else "no_compatible_backend",
            "providers": providers,
        }
    return {
        "id": gate_id,
        "kind": "unknown",
        "label": gate_id,
        "available": False,
        "reason": "unknown_capability_gate",
    }


def _route_id(plugin_id: str, route: RouteSpec) -> str:
    operation = route.action_name or "convert"
    return f"{plugin_id}:{route.source_format}:{route.target_format}:{operation}"


def _matching_contract(
    manifest: PluginManifest,
    route: RouteSpec,
) -> tuple[list[str], list[str], list[str], list[str]]:
    platforms = set(manifest.platforms)
    required: set[str] = set()
    optional: set[str] = set()
    limitations: set[str] = set()
    for rule in manifest.capability_rules:
        if not rule.matches(route):
            continue
        if rule.platforms:
            platforms.intersection_update(rule.platforms)
        required.update(rule.required_capabilities)
        optional.update(rule.optional_capabilities)
        limitations.update(rule.limitations)
    optional.difference_update(required)
    return sorted(platforms), sorted(required), sorted(optional), sorted(limitations)


def _project_route(
    manifest: PluginManifest,
    route: RouteSpec,
    *,
    platform_id: str,
    gates: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    platforms, required, optional, declared_limitations = _matching_contract(manifest, route)
    missing_required = [gate_id for gate_id in required if not bool(gates[gate_id]["available"])]
    missing_optional = [gate_id for gate_id in optional if not bool(gates[gate_id]["available"])]
    platform_supported = platform_id in platforms
    available = platform_supported and not missing_required

    limitations = list(declared_limitations)
    limitations.extend(f"optional_capability_unavailable:{gate_id}" for gate_id in missing_optional)
    if not platform_supported:
        limitations.append(f"platform_unsupported:{platform_id}")
    limitations.extend(f"required_capability_unavailable:{gate_id}" for gate_id in missing_required)

    if not available:
        state = "unavailable"
    elif limitations:
        state = "available_with_limits"
    else:
        state = "available"

    return {
        "id": _route_id(manifest.plugin_id, route),
        "operation": "action" if route.action_name else "conversion",
        "source": route.source_format,
        "target": route.target_format,
        "action": route.action_name or None,
        "label": route.label,
        "plugin": manifest.plugin_id,
        "available": available,
        "state": state,
        "platforms": platforms,
        "platform_supported": platform_supported,
        "required_capabilities": required,
        "optional_capabilities": optional,
        "missing_required_capabilities": missing_required,
        "missing_optional_capabilities": missing_optional,
        "limitations": sorted(set(limitations)),
        "options": sorted(str(key) for key in route.options_schema.get("properties", {})),
    }


def _optimization_scope(route: RouteSpec) -> tuple[str, str]:
    source_category = FORMAT_CATEGORY.get(route.source_format)
    if source_category is None and route.source_format in ALL_CATEGORIES:
        source_category = route.source_format
    if source_category is None:
        raise OptimizationContractError(f"Optimization route has no canonical source category: {route.source_format}")
    target = route.target_format.strip()
    if not target:
        raise OptimizationContractError("Optimization route target format must not be empty.")
    return f"{source_category}_to_{target}", source_category


def _project_optimizations(
    manifests: list[PluginManifest],
    *,
    platform_id: str,
    projected_routes: list[dict[str, Any]],
) -> dict[str, Any]:
    route_projection_by_id: dict[str, dict[str, Any]] = {}
    for projected_route in projected_routes:
        route_id = str(projected_route["id"])
        if route_id in route_projection_by_id:
            raise OptimizationContractError(f"Duplicate runtime route id: {route_id}")
        route_projection_by_id[route_id] = projected_route

    seen_resource_ids: set[str] = set()
    claimed_actions: set[tuple[str, str]] = set()
    resources: list[dict[str, Any]] = []
    for manifest in sorted(manifests, key=lambda item: item.plugin_id):
        for declaration in sorted(manifest.optimization_resources, key=lambda item: item.id):
            if declaration.id in seen_resource_ids:
                raise OptimizationContractError(f"Duplicate optimization resource id: {declaration.id}")
            seen_resource_ids.add(declaration.id)

            action_key = (manifest.plugin_id, declaration.action_name)
            if action_key in claimed_actions:
                raise OptimizationContractError(
                    f"Optimization action is declared by more than one resource: "
                    f"{manifest.plugin_id}/{declaration.action_name}"
                )
            claimed_actions.add(action_key)

            matching_routes = [route for route in manifest.routes if route.action_name == declaration.action_name]
            if not matching_routes:
                raise OptimizationContractError(
                    f"Optimization resource {declaration.id} has no matching action route: {declaration.action_name}"
                )

            bindings: list[dict[str, Any]] = []
            bound_scopes: set[str] = set()
            for route in matching_routes:
                scope, source_category = _optimization_scope(route)
                route_id = _route_id(manifest.plugin_id, route)
                projected_route = route_projection_by_id.get(route_id)
                if projected_route is None or projected_route.get("operation") != "action":
                    raise OptimizationContractError(
                        f"Optimization resource {declaration.id} references an unavailable route projection: {route_id}"
                    )
                bound_scopes.add(scope)
                bindings.append(
                    {
                        "scope": scope,
                        "route_id": route_id,
                        "source": route.source_format,
                        "source_category": source_category,
                        "target": route.target_format,
                        "available": bool(projected_route["available"]),
                        "state": str(projected_route["state"]),
                    }
                )

            bindings.sort(key=lambda item: (item["scope"], item["source"], item["target"], item["route_id"]))
            available_bindings = sum(bool(binding["available"]) for binding in bindings)
            if available_bindings == 0:
                state = "unavailable"
            elif available_bindings != len(bindings) or any(binding["state"] != "available" for binding in bindings):
                state = "available_with_limits"
            else:
                state = "available"
            resources.append(
                {
                    "id": declaration.id,
                    "name": declaration.name,
                    "action_name": declaration.action_name,
                    "scopes": sorted(bound_scopes),
                    "available": available_bindings > 0,
                    "state": state,
                    "bindings": bindings,
                }
            )

    resources.sort(key=lambda item: item["id"])
    bindings = [binding for resource in resources for binding in resource["bindings"]]
    return {
        "resource": "optimizations",
        "contract": {"id": OPTIMIZATION_CONTRACT_ID, "version": OPTIMIZATION_CONTRACT_VERSION},
        "runtime": {"state": "available", "platform": platform_id},
        "resources": resources,
        "counts": {
            "resources": len(resources),
            "available_resources": sum(bool(resource["available"]) for resource in resources),
            "unavailable_resources": sum(not bool(resource["available"]) for resource in resources),
            "bindings": len(bindings),
            "available_bindings": sum(bool(binding["available"]) for binding in bindings),
            "unavailable_bindings": sum(not bool(binding["available"]) for binding in bindings),
        },
    }


def build_runtime_capability_projection(
    manifests: list[PluginManifest],
    *,
    platform_id: str | None = None,
    egress_guard_status: Mapping[str, object] | None = None,
) -> dict[str, Any]:
    """Build the protocol-facing capability matrix for one loaded runtime.

    An initialized runtime with no registered routes is a successful empty
    result.  Failure to initialize or reach this projector is handled by the
    caller as a typed capability failure; it is never folded into an empty
    matrix.
    """

    platform_value = platform_id or current_platform_id()
    if egress_guard_status is None:
        from docwen_runtime.security.network import dependency_egress_guard_status

        guard_projection = dependency_egress_guard_status().to_dict()
    else:
        guard_projection = dict(egress_guard_status)
    local_transports = guard_projection.get("local_transports")
    if isinstance(local_transports, tuple):
        guard_projection["local_transports"] = list(local_transports)
    route_specs: list[tuple[PluginManifest, RouteSpec]] = []
    referenced_gates: set[str] = set()
    for manifest in sorted(manifests, key=lambda item: item.plugin_id):
        for route in manifest.routes:
            route_specs.append((manifest, route))
            _platforms, required, optional, _limitations = _matching_contract(manifest, route)
            referenced_gates.update(required)
            referenced_gates.update(optional)

    gate_map = {gate_id: _gate_status(gate_id) for gate_id in sorted(referenced_gates)}
    projected_routes = [
        _project_route(manifest, route, platform_id=platform_value, gates=gate_map) for manifest, route in route_specs
    ]
    projected_routes.sort(
        key=lambda item: (
            str(item["source"]),
            str(item["operation"]),
            str(item["action"] or ""),
            str(item["target"]),
            str(item["plugin"]),
        )
    )

    sources: list[dict[str, Any]] = []
    for source_id in sorted({str(item["source"]) for item in projected_routes}):
        source_routes = [item for item in projected_routes if item["source"] == source_id]
        sources.append(
            {
                "id": source_id,
                "category": FORMAT_CATEGORY.get(source_id, source_id),
                "available": any(bool(item["available"]) for item in source_routes),
                "routes": source_routes,
            }
        )

    return {
        "resource": "formats",
        "contract": {"id": CAPABILITY_CONTRACT_ID, "version": CAPABILITY_CONTRACT_VERSION},
        "runtime": {"state": "available", "platform": platform_value},
        "security": {"dependency_egress_guard": guard_projection},
        "gates": [gate_map[gate_id] for gate_id in sorted(gate_map)],
        "sources": sources,
        "counts": {
            "sources": len(sources),
            "routes": len(projected_routes),
            "available_routes": sum(bool(item["available"]) for item in projected_routes),
            "unavailable_routes": sum(not bool(item["available"]) for item in projected_routes),
            "actions": sum(item["operation"] == "action" for item in projected_routes),
        },
        "optimizations": _project_optimizations(
            manifests,
            platform_id=platform_value,
            projected_routes=projected_routes,
        ),
    }


__all__ = [
    "CAPABILITY_CONTRACT_ID",
    "CAPABILITY_CONTRACT_VERSION",
    "OPTIMIZATION_CONTRACT_ID",
    "OPTIMIZATION_CONTRACT_VERSION",
    "OptimizationContractError",
    "build_runtime_capability_projection",
    "current_platform_id",
    "probe_pymupdf_layout_resources",
]
