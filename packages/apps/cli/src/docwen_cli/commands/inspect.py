"""Inspect file metadata and project supported commands from Runtime routes."""

from __future__ import annotations

import argparse
import os
import sys
from typing import Any

from docwen_application.controller import CapabilityUnavailableError
from docwen_cli.capabilities import runtime_route_catalog
from docwen_cli.commands.execution_routes import STATIC_EXECUTION_ROUTES
from docwen_cli.exit_codes import ExitCode
from docwen_cli.file_admission_i18n import render_file_inspection_message
from docwen_cli.i18n import cli_t
from docwen_cli.parser import get_common_parser
from docwen_core.detection import inspect_file


def register_inspect_parser(subparsers: Any) -> argparse.ArgumentParser:
    """Register the ``inspect`` command."""
    p = subparsers.add_parser(
        "inspect",
        parents=[get_common_parser()],
        help=cli_t("cli.help.inspect"),
    )
    p.add_argument("file", help=cli_t("cli.help.files"))
    return p


def execute_inspect(args: argparse.Namespace, controller: Any = None) -> int:
    """Execute the inspect command.

    Returns 0 on success, 2 if the file cannot be inspected.
    """
    file_path = getattr(args, "file", "")
    if not file_path:
        return _print_inspect_error(args, "未指定文件")

    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        return _print_inspect_error(args, f"文件不存在: {file_path}", error_code="file_not_found")

    try:
        inspect_data = build_inspection_data(controller, abs_path)
    except FileNotFoundError:
        return _print_inspect_error(args, f"文件不存在: {file_path}", error_code="file_not_found")
    except OSError as exc:
        return _print_inspect_error(args, f"无法读取文件: {exc}")

    json_mode = bool(getattr(args, "json", False))
    if json_mode:
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter()
        presenter.present_data("inspect", inspect_data, success=True)
    else:
        _print_inspection_text(inspect_data)
    return int(ExitCode.OK)


def build_inspection_data(controller: Any, file_path: str) -> dict[str, Any]:
    """Build the canonical machine/CLI inspection payload without presenting it."""

    inspection = inspect_file(file_path)
    if inspection.may_execute:
        try:
            supported_actions, supported_actions_discovery = _discover_supported_actions(
                controller,
                inspection.detected_format,
                inspection.workflow_category,
            )
        except CapabilityUnavailableError as exc:
            supported_actions = ["inspect"]
            supported_actions_discovery = {
                "state": "unavailable",
                "matched_by": "none",
                "source_ids": [],
                "error": {"code": "capability_unavailable", "message": str(exc)},
            }
    else:
        supported_actions = ["inspect"]
        supported_actions_discovery = {
            "state": "not_applicable",
            "matched_by": "admission",
            "source_ids": [],
            "error": None,
        }

    inspect_data = inspection.to_dict()
    inspect_data["supported_actions"] = supported_actions
    inspect_data["supported_actions_discovery"] = supported_actions_discovery

    return inspect_data


def _print_inspection_text(inspect_data: dict[str, Any]) -> None:
    print(f"文件：{inspect_data['file_path']}")
    print(f"扩展名类别：{inspect_data['declared_category']}")
    print(f"扩展名格式：{inspect_data['declared_format']}")
    print(f"实际类别：{inspect_data['detected_category']}")
    print(f"实际格式：{inspect_data['detected_format']}")
    print(f"工作流类别：{inspect_data['workflow_category']}")
    print(f"检测方式：{inspect_data['detection_method']}")
    print(f"置信度：{inspect_data['confidence']}")
    print(f"格式关系：{inspect_data['relation']}")
    print(f"准入决定：{inspect_data['decision']}")
    primary_code = str(inspect_data["warning_code"] or inspect_data["reason_code"])
    if inspect_data["warning_code"]:
        print(f"⚠ [{inspect_data['warning_code']}] {render_file_inspection_message(_InspectionView(inspect_data))}")
    elif inspect_data["reason_code"]:
        print(
            f"阻断原因：[{inspect_data['reason_code']}] "
            f"{render_file_inspection_message(_InspectionView(inspect_data), prefer_reason=True)}"
        )
    for warning in inspect_data["warnings"]:
        code = str(warning.get("code", "")).strip()
        if code and code == primary_code:
            continue
        message = str(warning.get("message", "")).strip()
        label = f"[{code}] " if code else ""
        print(f"WARNING: {label}{message}")
    supported_actions_discovery = inspect_data["supported_actions_discovery"]
    if supported_actions_discovery["state"] == "unavailable":
        error = supported_actions_discovery["error"]
        if isinstance(error, dict):
            print(f"操作发现不可用：[{error['code']}] {error['message']}")
    print(f"支持的操作：{', '.join(inspect_data['supported_actions'])}")


class _InspectionView:
    """Attribute projection used by the CLI localization adapter."""

    def __init__(self, data: dict[str, Any]) -> None:
        self.declared_format = str(data.get("declared_format", "unknown"))
        self.detected_format = str(data.get("detected_format", "unknown"))
        self.warning_code = str(data.get("warning_code", ""))
        self.warning_message = str(data.get("warning_message", ""))
        self.reason_code = str(data.get("reason_code", ""))
        self.reason_message = str(data.get("reason_message", ""))
        self.warnings = tuple(item for item in data.get("warnings", []) if isinstance(item, dict))


def _discover_supported_actions(
    controller: Any | None,
    detected_format: str,
    workflow_category: str,
) -> tuple[list[str], dict[str, Any]]:
    """Join canonical Runtime routes to parser-owned public command paths."""

    catalog = runtime_route_catalog(controller)
    commands: list[str] = ["inspect"]
    public_paths_by_action: dict[str, list[str]] = {}
    for public_path, route in STATIC_EXECUTION_ROUTES.items():
        if public_path.startswith("batch "):
            continue
        public_paths_by_action.setdefault(route.action, []).append(public_path)

    for route in catalog.effective_routes(detected_format, workflow_category):
        if not route.available:
            continue
        if route.operation == "conversion":
            commands.append("convert")
            continue
        commands.extend(public_paths_by_action.get(route.action_name, ()))

    source_ids = catalog.source_ids_for(detected_format, workflow_category)
    first_source = next(iter(source_ids), None)
    if first_source is None:
        matched_by = "none"
    elif first_source == detected_format and len(source_ids) > 1:
        matched_by = "detected_format_then_workflow_category"
    elif first_source == detected_format:
        matched_by = "detected_format"
    else:
        matched_by = "workflow_category"
    return list(dict.fromkeys(commands)), {
        "state": "available",
        "matched_by": matched_by,
        "source_ids": list(source_ids),
        "error": None,
    }


def _print_inspect_error(args: argparse.Namespace, message: str, *, error_code: str = "invalid_input") -> int:
    """Print an inspect error and return exit code 2."""
    if getattr(args, "json", False):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter()
        presenter.present_error("inspect", message, error_code=error_code)
    else:
        print(f"错误: {message}", file=sys.stderr)
    return int(ExitCode.INVALID_INPUT)
