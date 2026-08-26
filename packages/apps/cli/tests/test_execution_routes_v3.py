"""Canonical public-command to runtime-action mapping."""

from __future__ import annotations

import argparse

import pytest

pytestmark = pytest.mark.contract


def test_every_static_route_has_one_public_path_and_action() -> None:
    from docwen_cli.commands.execution_routes import STATIC_EXECUTION_ROUTES

    assert set(STATIC_EXECUTION_ROUTES) == {
        "validate",
        "number markdown",
        "merge pdf",
        "merge tables",
        "merge images",
        "split pdf",
        "batch validate",
    }
    assert all(path == route.public_path for path, route in STATIC_EXECUTION_ROUTES.items())


def test_convert_route_does_not_guess_action_from_public_optimization_id() -> None:
    from docwen_cli.commands.execution_routes import route_for_public_command

    route = route_for_public_command("convert")
    assert route.action == ""
    assert not hasattr(route, "target_format")


def test_execution_resolves_public_optimization_id_through_runtime_projection(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_cli.commands import execution_v3

    runtime = {"state": "available", "platform": "windows"}

    class Controller:
        def describe_runtime_capabilities(self) -> dict[str, object]:
            return {
                "resource": "formats",
                "contract": {"id": "docwen.runtime-capabilities", "version": 1},
                "runtime": runtime,
                "security": {"dependency_egress_guard": {}},
                "gates": [],
                "sources": [
                    {
                        "id": "docx",
                        "category": "document",
                        "available": True,
                        "routes": [
                            {
                                "id": "optimizer:docx:md:internal_action",
                                "operation": "action",
                                "source": "docx",
                                "target": "md",
                                "action": "internal_action",
                                "available": True,
                                "state": "available",
                                "options": [],
                            }
                        ],
                    }
                ],
                "counts": {
                    "sources": 1,
                    "routes": 1,
                    "available_routes": 1,
                    "unavailable_routes": 0,
                    "actions": 1,
                },
                "optimizations": {
                    "resource": "optimizations",
                    "contract": {"id": "docwen.optimizations", "version": 1},
                    "runtime": runtime,
                    "resources": [
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
                    ],
                    "counts": {
                        "resources": 1,
                        "available_resources": 1,
                        "unavailable_resources": 0,
                        "bindings": 1,
                        "available_bindings": 1,
                        "unavailable_bindings": 0,
                    },
                },
            }

    captured: dict[str, str] = {}

    def fake_execute(args: argparse.Namespace, _controller: object) -> int:
        captured["action"] = str(args.action)
        return 0

    monkeypatch.setattr(execution_v3, "execute_convert", fake_execute)
    args = argparse.Namespace(
        _docwen_preflight_done=True,
        optimization="public-resource",
        action="",
    )

    assert execution_v3.execute_execution(args, Controller()) == 0
    assert captured == {"action": "internal_action"}
