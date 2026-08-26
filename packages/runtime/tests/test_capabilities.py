"""Runtime composition capability projection contracts."""

from __future__ import annotations

from typing import Any

import pytest

from docwen_core.models.manifest import OptimizationResourceSpec, PluginManifest, RouteCapabilityRule, RouteSpec
from docwen_runtime import capabilities

pytestmark = pytest.mark.unit


def _status(gate_id: str, *, available: bool) -> dict[str, Any]:
    return {
        "id": gate_id,
        "kind": "test",
        "label": gate_id,
        "available": available,
        "reason": None if available else "test_missing",
    }


def test_projection_preserves_actions_and_route_level_gate_failures(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    statuses = {
        "test.optional": _status("test.optional", available=False),
        "test.required": _status("test.required", available=False),
    }
    monkeypatch.setattr(capabilities, "_gate_status", lambda gate_id: statuses[gate_id])
    manifest = PluginManifest(
        plugin_id="probe",
        name="Probe",
        version="1",
        routes=[
            RouteSpec("alpha", "beta", label="Alpha to beta", options_schema={"properties": {"flag": {}}}),
            RouteSpec("alpha", "alpha", action_name="validate", label="Validate alpha"),
        ],
        capability_rules=[
            RouteCapabilityRule(
                operations=("conversion",),
                optional_capabilities=("test.optional",),
                limitations=("conversion_limit",),
            ),
            RouteCapabilityRule(
                action_names=("validate",),
                required_capabilities=("test.required",),
            ),
        ],
    )

    projection = capabilities.build_runtime_capability_projection([manifest], platform_id="windows")

    assert projection["contract"] == {"id": "docwen.runtime-capabilities", "version": 1}
    assert projection["counts"] == {
        "sources": 1,
        "routes": 2,
        "available_routes": 1,
        "unavailable_routes": 1,
        "actions": 1,
    }
    routes = projection["sources"][0]["routes"]
    conversion = next(item for item in routes if item["operation"] == "conversion")
    action = next(item for item in routes if item["operation"] == "action")
    assert conversion["available"] is True
    assert conversion["state"] == "available_with_limits"
    assert conversion["missing_optional_capabilities"] == ["test.optional"]
    assert conversion["options"] == ["flag"]
    assert action["action"] == "validate"
    assert action["available"] is False
    assert action["state"] == "unavailable"
    assert action["missing_required_capabilities"] == ["test.required"]
    assert projection["optimizations"]["resources"] == []


def test_projection_fails_closed_for_unknown_declared_gate() -> None:
    manifest = PluginManifest(
        plugin_id="probe",
        name="Probe",
        version="1",
        routes=[RouteSpec("alpha", "beta")],
        capability_rules=[RouteCapabilityRule(required_capabilities=("unknown.gate",))],
    )

    projection = capabilities.build_runtime_capability_projection([manifest], platform_id="windows")

    gate = projection["gates"][0]
    route = projection["sources"][0]["routes"][0]
    assert gate["available"] is False
    assert gate["reason"] == "unknown_capability_gate"
    assert route["available"] is False
    assert route["limitations"] == ["required_capability_unavailable:unknown.gate"]


def test_projection_marks_unsupported_platform_unavailable(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(capabilities, "_gate_status", lambda gate_id: _status(gate_id, available=True))
    manifest = PluginManifest(
        plugin_id="probe",
        name="Probe",
        version="1",
        routes=[RouteSpec("alpha", "beta")],
        platforms=("windows",),
    )

    projection = capabilities.build_runtime_capability_projection([manifest], platform_id="linux")

    route = projection["sources"][0]["routes"][0]
    assert route["platform_supported"] is False
    assert route["available"] is False
    assert route["limitations"] == ["platform_unsupported:linux"]


def test_initialized_runtime_with_no_routes_is_successful_empty_projection() -> None:
    projection = capabilities.build_runtime_capability_projection(
        [],
        platform_id="windows",
        egress_guard_status={
            "state": "enforced",
            "installed": True,
            "active": True,
            "scope": "docwen_python_process",
            "policy": "deny_dns_and_ip",
            "mechanism": "cpython_audit_hook",
            "bootstrap": "composition_root",
            "local_transports": ("windows_named_pipe", "unix_domain_socket"),
            "external_processes": "not_managed",
        },
    )

    assert projection["runtime"] == {"state": "available", "platform": "windows"}
    assert projection["security"]["dependency_egress_guard"] == {
        "state": "enforced",
        "installed": True,
        "active": True,
        "scope": "docwen_python_process",
        "policy": "deny_dns_and_ip",
        "mechanism": "cpython_audit_hook",
        "bootstrap": "composition_root",
        "local_transports": ["windows_named_pipe", "unix_domain_socket"],
        "external_processes": "not_managed",
    }
    assert projection["gates"] == []
    assert projection["sources"] == []
    assert projection["counts"] == {
        "sources": 0,
        "routes": 0,
        "available_routes": 0,
        "unavailable_routes": 0,
        "actions": 0,
    }
    assert projection["optimizations"] == {
        "resource": "optimizations",
        "contract": {"id": "docwen.optimizations", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
        "resources": [],
        "counts": {
            "resources": 0,
            "available_resources": 0,
            "unavailable_resources": 0,
            "bindings": 0,
            "available_bindings": 0,
            "unavailable_bindings": 0,
        },
    }


def test_optimization_projection_binds_public_id_to_independent_action_and_routes() -> None:
    manifest = PluginManifest(
        plugin_id="optimizer",
        name="Optimizer",
        version="1",
        routes=[
            RouteSpec("docx", "md", action_name="internal_action"),
        ],
        optimization_resources=[
            OptimizationResourceSpec(
                id="public-resource",
                name="Public resource",
                action_name="internal_action",
            )
        ],
    )

    projection = capabilities.build_runtime_capability_projection([manifest], platform_id="windows")

    optimizations = projection["optimizations"]
    assert optimizations["contract"] == {"id": "docwen.optimizations", "version": 1}
    assert optimizations["counts"] == {
        "resources": 1,
        "available_resources": 1,
        "unavailable_resources": 0,
        "bindings": 1,
        "available_bindings": 1,
        "unavailable_bindings": 0,
    }
    assert optimizations["resources"] == [
        {
            "id": "public-resource",
            "name": "Public resource",
            "action_name": "internal_action",
            "scopes": ["document_to_md"],
            "available": True,
            "state": "available",
            "bindings": [
                {
                    "scope": "document_to_md",
                    "route_id": "optimizer:docx:md:internal_action",
                    "source": "docx",
                    "source_category": "document",
                    "target": "md",
                    "available": True,
                    "state": "available",
                }
            ],
        }
    ]


@pytest.mark.parametrize(
    ("routes", "resources", "message"),
    [
        (
            [],
            [OptimizationResourceSpec("resource", "Resource", "action")],
            "no matching action route",
        ),
        (
            [RouteSpec("mystery", "md", action_name="action")],
            [OptimizationResourceSpec("resource", "Resource", "action")],
            "no canonical source category",
        ),
    ],
)
def test_optimization_projection_fails_closed_for_declaration_route_mismatch(
    routes: list[RouteSpec],
    resources: list[OptimizationResourceSpec],
    message: str,
) -> None:
    manifest = PluginManifest(
        plugin_id="optimizer",
        name="Optimizer",
        version="1",
        routes=routes,
        optimization_resources=resources,
    )

    with pytest.raises(capabilities.OptimizationContractError, match=message):
        capabilities.build_runtime_capability_projection([manifest], platform_id="windows")


def test_optimization_projection_rejects_duplicate_public_ids() -> None:
    manifests = [
        PluginManifest(
            plugin_id=f"optimizer-{index}",
            name="Optimizer",
            version="1",
            routes=[RouteSpec("docx", "md", action_name=f"action_{index}")],
            optimization_resources=[
                OptimizationResourceSpec(
                    "shared",
                    "Shared",
                    f"action_{index}",
                )
            ],
        )
        for index in range(2)
    ]

    with pytest.raises(capabilities.OptimizationContractError, match="Duplicate optimization resource id"):
        capabilities.build_runtime_capability_projection(manifests, platform_id="windows")


def test_bundled_optimizer_manifests_project_canonical_resources(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_plugin_optimizer_gongwen.manifest import build_manifest as build_gongwen_manifest
    from docwen_plugin_optimizer_invoice_cn.manifest import build_manifest as build_invoice_manifest

    monkeypatch.setattr(capabilities, "_gate_status", lambda gate_id: _status(gate_id, available=True))

    projection = capabilities.build_runtime_capability_projection(
        [build_gongwen_manifest(), build_invoice_manifest()],
        platform_id="windows",
    )["optimizations"]

    resources = {resource["id"]: resource for resource in projection["resources"]}
    assert set(resources) == {"gongwen", "invoice_cn"}
    assert resources["gongwen"]["scopes"] == ["document_to_md"]
    assert resources["invoice_cn"]["scopes"] == ["image_to_md", "layout_to_md"]
    assert {binding["source"] for binding in resources["invoice_cn"]["bindings"]} == {"pdf", "ofd", "image"}


def test_bundled_optimizer_manifests_do_not_gate_capabilities_by_ui_locale() -> None:
    from docwen_plugin_optimizer_gongwen.manifest import build_manifest as build_gongwen_manifest
    from docwen_plugin_optimizer_invoice_cn.manifest import build_manifest as build_invoice_manifest

    manifests = (build_gongwen_manifest(), build_invoice_manifest())

    assert all("locales" not in manifest.extra for manifest in manifests)


def test_pymupdf_layout_resource_probe_rejects_unrelated_yaml_and_onnx(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    package_root = tmp_path / "layout"
    resources = package_root / "resources"
    resources.mkdir(parents=True)
    (resources / "unrelated.onnx").write_bytes(b"x")
    (resources / "unrelated.yaml").write_bytes(b"x")
    monkeypatch.setattr(capabilities, "_module_available", lambda _module: True)
    monkeypatch.setattr(capabilities, "_pymupdf_layout_package_roots", lambda: (package_root,))

    status = capabilities.probe_pymupdf_layout_resources()

    assert status["available"] is False
    assert status["reason"] == "required_resource_missing"
    assert status["resource_types"] == []
    assert str(tmp_path) not in str(status)


def test_pymupdf_layout_resource_probe_verifies_the_pinned_installed_resources() -> None:
    status = capabilities.probe_pymupdf_layout_resources()

    assert status == {
        "available": True,
        "reason": None,
        "resource_types": ["onnx", "yaml"],
        "resource_count": 7,
    }


def test_pymupdf_capability_gate_fails_closed_when_resources_are_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        capabilities,
        "probe_pymupdf_layout_resources",
        lambda: {
            "available": False,
            "reason": "required_resource_missing",
            "resource_types": ["yaml"],
            "resource_count": 1,
        },
    )

    status = capabilities._gate_status("python.pymupdf4llm")

    assert status["available"] is False
    assert status["reason"] == "required_resource_missing"
    assert status["kind"] == "python_module_with_resources"
