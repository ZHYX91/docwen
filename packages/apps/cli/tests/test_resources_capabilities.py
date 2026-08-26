"""Protocol 3 resource discovery must reflect the injected runtime."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.contract


def _optimization_projection(*, resources: list[dict[str, Any]]) -> dict[str, Any]:
    bindings = [binding for resource in resources for binding in resource["bindings"]]
    return {
        "resource": "optimizations",
        "contract": {"id": "docwen.optimizations", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
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


def _projection(
    *,
    sources: list[dict[str, Any]],
    optimizations: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    if optimizations and not sources:
        source_groups: dict[str, dict[str, Any]] = {}
        for resource in optimizations:
            for binding in resource["bindings"]:
                source = str(binding["source"])
                group = source_groups.setdefault(
                    source,
                    {
                        "id": source,
                        "category": binding["source_category"],
                        "available": bool(binding["available"]),
                        "routes": [],
                    },
                )
                group["routes"].append(
                    {
                        "id": binding["route_id"],
                        "operation": "action",
                        "source": source,
                        "target": binding["target"],
                        "action": resource["action_name"],
                        "available": binding["available"],
                        "state": binding["state"],
                        "options": [],
                    }
                )
        sources = list(source_groups.values())
    routes = [route for source in sources for route in source["routes"]]
    return {
        "resource": "formats",
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
        "security": {
            "dependency_egress_guard": {
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
        },
        "gates": [],
        "sources": sources,
        "counts": {
            "sources": len(sources),
            "routes": len(routes),
            "available_routes": sum(bool(route["available"]) for route in routes),
            "unavailable_routes": sum(not bool(route["available"]) for route in routes),
            "actions": sum(route["operation"] == "action" for route in routes),
        },
        "optimizations": _optimization_projection(resources=optimizations or []),
    }


class _DiscoveryController:
    def __init__(self, projection: dict[str, Any], *, config_port: Any | None = None) -> None:
        self._projection = projection
        self.config_port = config_port

    def describe_runtime_capabilities(self) -> dict[str, Any]:
        return self._projection


class _TemplateRegistryStub:
    def __init__(self, entries: list[Any]) -> None:
        self._entries = entries

    def list_templates(self, target: str | None = None) -> list[Any]:
        if target is None:
            return list(self._entries)
        return [entry for entry in self._entries if entry.target == target]


class _ConfigPortStub:
    def __init__(self, snapshot: dict[str, Any]) -> None:
        self._snapshot = snapshot

    def snapshot(self) -> dict[str, Any]:
        return self._snapshot


def test_template_resource_id_is_listed_and_show_uses_the_same_token(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main
    from docwen_runtime.templates import TemplateInfo, TemplateRegistry

    template_id = "template.docx.4fd9bdd9a72279293b810865533f9662a4668e281a891e43d4d6f5939adc5c09"
    template = TemplateInfo(
        id=template_id,
        name="Corporate Report",
        target="docx",
        description="Corporate Report DOCX template",
        path=Path("D:/DocWen/templates/Corporate Report.docx"),
        size_bytes=1234,
        modified_ns=5678,
    )
    registry = _TemplateRegistryStub([template])
    monkeypatch.setattr(TemplateRegistry, "default", classmethod(lambda cls, extra_paths=None: registry))
    controller = _DiscoveryController(_projection(sources=[]))

    assert main(["resources", "list", "templates", "--json", "--quiet"], controller=controller) == 0
    listed = json.loads(capsys.readouterr().out)
    assert listed["data"] == {"type": "templates", "resources": [template.to_dict()], "total": 1}

    assert main(["resources", "show", "templates", template_id, "--json", "--quiet"], controller=controller) == 0
    shown = json.loads(capsys.readouterr().out)
    assert shown["data"] == {"type": "templates", "resource": template.to_dict()}


def test_template_identity_conflict_is_typed_discovery_failure(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main
    from docwen_runtime.templates import TemplateIdentityConflictError, TemplateRegistry

    class ConflictingRegistry:
        @staticmethod
        def list_templates(target: str | None = None) -> list[Any]:
            del target
            raise TemplateIdentityConflictError("conflict")

    monkeypatch.setattr(
        TemplateRegistry,
        "default",
        classmethod(lambda cls, extra_paths=None: ConflictingRegistry()),
    )

    exit_code = main(
        ["resources", "list", "templates", "--json", "--quiet"],
        controller=_DiscoveryController(_projection(sources=[])),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 6
    assert payload["success"] is False
    assert payload["error"]["code"] == "capability_unavailable"


def test_numbering_resources_filter_locale_and_accept_wildcard(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main
    from docwen_runtime.resources import ResourceRegistry

    snapshot = {
        "gui": {"language": {"locale": "zh_CN"}},
        "numbering": {
            "add": {
                "settings": {"order": ["wildcard", "en_only", "unrestricted"]},
                "schemes": {
                    "wildcard": {
                        "name": "Wildcard",
                        "locales": ["*"],
                        "level_1": {"format": "{n}."},
                    },
                    "en_only": {
                        "name": "English only",
                        "locales": ["en_US"],
                        "level_1": {"format": "{n}."},
                    },
                    "unrestricted": {
                        "name": "Unrestricted",
                        "level_1": {"format": "{n}."},
                    },
                },
            }
        },
    }
    controller = _DiscoveryController(
        _projection(sources=[]),
        config_port=_ConfigPortStub(snapshot),
    )
    monkeypatch.setattr(
        ResourceRegistry,
        "default",
        classmethod(lambda cls: type("_Resources", (), {"locales_dir": lambda self: tmp_path})()),
    )

    assert (
        main(
            ["resources", "list", "numbering-schemes", "--json", "--quiet"],
            controller=controller,
        )
        == 0
    )
    zh_payload = json.loads(capsys.readouterr().out)
    assert [item["id"] for item in zh_payload["data"]["resources"]] == [
        "wildcard",
        "unrestricted",
    ]

    assert (
        main(
            ["resources", "list", "numbering-schemes", "--lang", "en_US", "--json", "--quiet"],
            controller=controller,
        )
        == 0
    )
    en_payload = json.loads(capsys.readouterr().out)
    assert [item["id"] for item in en_payload["data"]["resources"]] == [
        "wildcard",
        "en_only",
        "unrestricted",
    ]

    assert (
        main(
            ["resources", "show", "numbering-schemes", "wildcard", "--json", "--quiet"],
            controller=controller,
        )
        == 0
    )
    shown_payload = json.loads(capsys.readouterr().out)
    assert shown_payload["data"]["resource"]["id"] == "wildcard"

    assert (
        main(
            ["resources", "show", "numbering-schemes", "Wildcard", "--json", "--quiet"],
            controller=controller,
        )
        == 3
    )
    rejected_payload = json.loads(capsys.readouterr().out)
    assert rejected_payload["error"]["code"] == "resource_not_found"


def test_resources_formats_emits_runtime_matrix_and_action_route(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    action = {
        "id": "probe:pdf:pdf:split_pdf",
        "operation": "action",
        "source": "pdf",
        "target": "pdf",
        "action": "split_pdf",
        "plugin": "probe",
        "available": False,
        "state": "unavailable",
        "platforms": ["windows"],
        "platform_supported": True,
        "required_capabilities": ["python.fitz"],
        "optional_capabilities": [],
        "missing_required_capabilities": ["python.fitz"],
        "missing_optional_capabilities": [],
        "limitations": ["required_capability_unavailable:python.fitz"],
        "options": [],
    }
    source = {"id": "pdf", "category": "layout", "available": False, "routes": [action]}

    exit_code = main(
        ["resources", "list", "formats", "--json", "--quiet"],
        controller=_DiscoveryController(_projection(sources=[source])),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["success"] is True
    assert payload["data"]["contract"] == {"id": "docwen.runtime-capabilities", "version": 1}
    assert payload["data"]["sources"][0]["routes"][0]["action"] == "split_pdf"
    assert payload["data"]["sources"][0]["routes"][0]["available"] is False


def test_resources_formats_successfully_preserves_empty_runtime(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    exit_code = main(
        ["resources", "list", "formats", "--json"],
        controller=_DiscoveryController(_projection(sources=[])),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["success"] is True
    assert payload["data"]["sources"] == []
    assert payload["data"]["counts"]["routes"] == 0


def test_resources_formats_without_runtime_is_typed_failure(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    exit_code = main(["resources", "list", "formats", "--json", "--quiet"])

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 6
    assert payload["success"] is False
    assert payload["data"] == {}
    assert payload["error"]["category"] == "unavailable"
    assert payload["error"]["code"] == "capability_unavailable"


def test_resources_formats_without_egress_guard_status_is_typed_failure(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    projection = _projection(sources=[])
    projection.pop("security")
    exit_code = main(
        ["resources", "list", "formats", "--json", "--quiet"],
        controller=_DiscoveryController(projection),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 6
    assert payload["success"] is False
    assert payload["error"]["code"] == "capability_unavailable"


def test_resources_show_formats_selects_source(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    source = {"id": "markdown", "category": "markdown", "available": True, "routes": []}
    exit_code = main(
        ["resources", "show", "formats", "markdown", "--json"],
        controller=_DiscoveryController(_projection(sources=[source])),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["data"]["source"] == source
    assert "sources" not in payload["data"]


def _resource() -> dict[str, Any]:
    return {
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


def test_resources_optimizations_lists_runtime_projection_without_guessing_action(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    resource = _resource()
    exit_code = main(
        ["resources", "list", "optimizations", "--json", "--quiet"],
        controller=_DiscoveryController(_projection(sources=[], optimizations=[resource])),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["data"] == _optimization_projection(resources=[resource])
    assert payload["data"]["resources"][0]["id"] != payload["data"]["resources"][0]["action_name"]


def test_resources_optimizations_inventory_is_locale_invariant(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    resource = _resource()
    controller = _DiscoveryController(_projection(sources=[], optimizations=[resource]))
    outputs: dict[str, str] = {}
    for locale in ("en_US", "zh_CN"):
        assert (
            main(
                ["resources", "list", "optimizations", "--lang", locale, "--json", "--quiet"],
                controller=controller,
            )
            == 0
        )
        outputs[locale] = capsys.readouterr().out

    assert outputs["en_US"] == outputs["zh_CN"]
    payload = json.loads(outputs["en_US"])
    assert [item["id"] for item in payload["data"]["resources"]] == ["public-resource"]


def test_resources_show_optimization_preserves_contract_and_binding(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    resource = _resource()
    exit_code = main(
        ["resources", "show", "optimizations", "public-resource", "--json"],
        controller=_DiscoveryController(_projection(sources=[], optimizations=[resource])),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["data"] == {
        "resource": "optimizations",
        "contract": {"id": "docwen.optimizations", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
        "optimization": resource,
    }


def test_resources_optimizations_successfully_preserves_empty_runtime(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    exit_code = main(
        ["resources", "list", "optimizations", "--json", "--quiet"],
        controller=_DiscoveryController(_projection(sources=[])),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["data"] == _optimization_projection(resources=[])


@pytest.mark.parametrize("mutation", ["missing", "bad_binding", "bad_counts"])
def test_resources_optimizations_malformed_projection_is_typed_failure(
    mutation: str,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    projection = _projection(sources=[], optimizations=[_resource()])
    if mutation == "missing":
        projection.pop("optimizations")
    elif mutation == "bad_binding":
        projection["optimizations"]["resources"][0]["bindings"][0].pop("route_id")
    else:
        projection["optimizations"]["counts"]["bindings"] = 99

    exit_code = main(
        ["resources", "list", "optimizations", "--json", "--quiet"],
        controller=_DiscoveryController(projection),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 6
    assert payload["error"]["code"] == "capability_unavailable"


def test_optimization_resource_id_resolves_declared_action() -> None:
    from docwen_cli.capabilities import resolve_optimization_action

    controller = _DiscoveryController(_projection(sources=[], optimizations=[_resource()]))

    assert resolve_optimization_action(controller, "public-resource") == "internal_action"
    with pytest.raises(ValueError, match="Unknown optimization resource"):
        resolve_optimization_action(controller, "internal_action")
