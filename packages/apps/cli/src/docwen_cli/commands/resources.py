"""Unified protocol 3 resource discovery."""

from __future__ import annotations

import argparse
import sys
from typing import Any

from docwen_application.controller import CapabilityUnavailableError
from docwen_cli.capabilities import runtime_capability_projection, runtime_optimization_projection
from docwen_cli.exit_codes import ExitCode
from docwen_cli.parser import get_common_parser
from docwen_cli.presenters.json_presenter import JsonPresenter

RESOURCE_TYPES = ("formats", "optimizations", "templates", "numbering-schemes")


def _format_projection(controller: Any | None) -> dict[str, Any]:
    return runtime_capability_projection(controller)


def _template_entries(args: argparse.Namespace) -> list[dict[str, Any]]:
    from docwen_runtime.templates import TemplateIdentityConflictError, TemplateRegistry

    try:
        templates = TemplateRegistry.default().list_templates(getattr(args, "target", None))
    except TemplateIdentityConflictError as exc:
        raise CapabilityUnavailableError("Template resource discovery found conflicting canonical IDs.") from exc
    return [item.to_dict() for item in templates]


def _numbering_entries(args: argparse.Namespace, controller: Any | None) -> list[dict[str, Any]]:
    from docwen_runtime.numbering import NumberingSchemeRegistry
    from docwen_runtime.resources import ResourceRegistry

    config_port = getattr(controller, "config_port", None)
    if config_port is None:
        raise CapabilityUnavailableError("Numbering resource discovery requires an injected configuration port.")
    try:
        snapshot = config_port.snapshot()
    except Exception as exc:
        raise CapabilityUnavailableError("Numbering configuration discovery failed.") from exc
    if not isinstance(snapshot, dict):
        raise CapabilityUnavailableError("Numbering configuration discovery returned an invalid snapshot.")
    locale = str(getattr(args, "lang", None) or _snapshot_locale(snapshot) or "zh_CN")
    registry = NumberingSchemeRegistry.from_config_snapshot(
        snapshot,
        locale_path=ResourceRegistry.default().locales_dir() / f"{locale}.toml",
    )
    return [item.to_dict() for item in registry.list_schemes()]


def _entries(resource_type: str, args: argparse.Namespace, controller: Any | None) -> list[dict[str, Any]]:
    if resource_type == "templates":
        return _template_entries(args)
    if resource_type == "numbering-schemes":
        return _numbering_entries(args, controller)
    raise ValueError(f"Unknown resource type: {resource_type}")


def execute_resources(args: argparse.Namespace, controller: Any | None = None) -> int:
    resource_type = str(args.resource_type)
    command = f"resources {args.resources_command}"
    if resource_type in {"formats", "optimizations"}:
        projection = (
            _format_projection(controller)
            if resource_type == "formats"
            else runtime_optimization_projection(controller)
        )
        if args.resources_command == "show":
            resource_id = str(args.resource_id)
            raw_entries = projection.get("sources" if resource_type == "formats" else "resources", [])
            if not isinstance(raw_entries, list):
                raise CapabilityUnavailableError("Runtime resource discovery returned an invalid projection.")
            entries = raw_entries
            match = next(
                (item for item in entries if isinstance(item, dict) and str(item.get("id", "")) == resource_id),
                None,
            )
            if match is None:
                message = f"Resource not found: {resource_type}/{resource_id}"
                if getattr(args, "json", False):
                    JsonPresenter().present_error(command, message, error_code="resource_not_found")
                else:
                    print(f"Error: {message}", file=sys.stderr)
                return int(ExitCode.NOT_FOUND)
            data = {
                "resource": resource_type,
                "contract": projection.get("contract"),
                "runtime": projection.get("runtime"),
            }
            if resource_type == "formats":
                data["gates"] = projection.get("gates", [])
                data["source"] = match
            else:
                data["optimization"] = match
        else:
            data = projection

        if getattr(args, "json", False):
            JsonPresenter(include_timing=getattr(args, "timing", False)).present_data(command, data)
        elif args.resources_command == "show":
            if resource_type == "formats":
                _print_format_source(data["source"])
            else:
                _print_optimization(data["optimization"])
        else:
            key = "sources" if resource_type == "formats" else "resources"
            for item in data.get(key, []):
                if isinstance(item, dict):
                    if resource_type == "formats":
                        _print_format_source(item)
                    else:
                        _print_optimization(item)
        return int(ExitCode.OK)

    entries = _entries(resource_type, args, controller)

    if args.resources_command == "show":
        resource_id = str(args.resource_id)
        match = next((item for item in entries if str(item.get("id", "")) == resource_id), None)
        if match is None:
            message = f"Resource not found: {resource_type}/{resource_id}"
            if getattr(args, "json", False):
                JsonPresenter().present_error(command, message, error_code="resource_not_found")
            else:
                print(f"Error: {message}", file=sys.stderr)
            return int(ExitCode.NOT_FOUND)
        data: dict[str, Any] = {"type": resource_type, "resource": match}
    else:
        data = {"type": resource_type, "resources": entries, "total": len(entries)}

    if getattr(args, "json", False):
        JsonPresenter(include_timing=getattr(args, "timing", False)).present_data(command, data)
    else:
        if args.resources_command == "show":
            for key, value in data["resource"].items():
                print(f"{key}: {value}")
        else:
            for item in entries:
                print(f"{item.get('id', item.get('name', ''))}\t{item.get('name', '')}")
    return int(ExitCode.OK)


def _snapshot_locale(snapshot: dict[str, Any]) -> str | None:
    gui = snapshot.get("gui", {})
    language = gui.get("language", {}) if isinstance(gui, dict) else {}
    locale = language.get("locale") if isinstance(language, dict) else None
    return str(locale).strip() if locale else None


def _print_format_source(source: dict[str, Any]) -> None:
    routes = source.get("routes", [])
    available = sum(bool(route.get("available")) for route in routes if isinstance(route, dict))
    total = sum(isinstance(route, dict) for route in routes)
    targets = sorted(
        {
            str(route.get("target"))
            for route in routes
            if isinstance(route, dict) and route.get("operation") == "conversion"
        }
    )
    actions = sorted(
        {str(route.get("action")) for route in routes if isinstance(route, dict) and route.get("operation") == "action"}
    )
    summary = f"{source.get('id', '')}\t{available}/{total} available"
    if targets:
        summary += f"\ttargets={','.join(targets)}"
    if actions:
        summary += f"\tactions={','.join(actions)}"
    print(summary)


def _print_optimization(resource: dict[str, Any]) -> None:
    scopes = ",".join(str(scope) for scope in resource.get("scopes", []))
    print(f"{resource.get('id', '')}\t{resource.get('name', '')}\tstate={resource.get('state', '')}\tscopes={scopes}")


def register_resources_parser(subparsers: Any) -> argparse.ArgumentParser:
    parser = subparsers.add_parser(
        "resources",
        parents=[get_common_parser()],
        help="List or inspect stable DocWen resources.",
    )
    commands = parser.add_subparsers(dest="resources_command", required=True)

    list_parser = commands.add_parser("list", parents=[get_common_parser()])
    list_parser.add_argument("resource_type", choices=RESOURCE_TYPES, metavar="TYPE")
    list_parser.add_argument("--target", choices=["docx", "xlsx"])

    show_parser = commands.add_parser("show", parents=[get_common_parser()])
    show_parser.add_argument("resource_type", choices=RESOURCE_TYPES, metavar="TYPE")
    show_parser.add_argument("resource_id", metavar="ID")
    return parser


__all__ = ["RESOURCE_TYPES", "execute_resources", "register_resources_parser"]
