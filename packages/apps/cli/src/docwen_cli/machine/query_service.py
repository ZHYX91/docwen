"""Typed read/control operations exposed by DocWen Machine Protocol v1."""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path
from typing import Any, Literal

from docwen_application.controller import CapabilityUnavailableError
from docwen_cli.capabilities import runtime_capability_projection, runtime_optimization_projection
from docwen_cli.commands.doctor import collect_doctor_checks
from docwen_cli.commands.inspect import build_inspection_data
from docwen_cli.commands.resources import _numbering_entries, _template_entries
from docwen_cli.gui_control_port import GuiControlError, GuiControlPort

_HASH_CHUNK_BYTES = 1024 * 1024
ResourceKind = Literal["formats", "optimizations", "templates", "numbering-schemes"]


class MachineQueryError(ValueError):
    """Stable query/control failure projected as JSON-RPC error data."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


class MachineQueryService:
    """Own non-task machine queries without invoking CLI parsers or presenters."""

    def __init__(self, controller: Any, gui_control: GuiControlPort | None = None) -> None:
        self._controller = controller
        self._gui_control = gui_control

    def health_check(self) -> dict[str, Any]:
        projection = runtime_capability_projection(self._controller)
        checks = collect_doctor_checks(self._controller, projection)
        return {
            "all_ok": all(check.ok for check in checks),
            "checks": [check.to_dict() for check in checks],
        }

    def inspect_file(self, input_handle: dict[str, Any]) -> dict[str, Any]:
        path = self._verified_file(input_handle)
        try:
            inspection = build_inspection_data(self._controller, str(path))
        except FileNotFoundError as exc:
            raise MachineQueryError("file_not_found", "The input file does not exist.") from exc
        except OSError as exc:
            raise MachineQueryError("file_unreadable", "The input file cannot be inspected.") from exc
        return {
            key: inspection.get(key)
            for key in (
                "file_path",
                "size_bytes",
                "content_sha256",
                "decision",
                "supported_actions",
                "declared_format",
                "detected_format",
                "warning_code",
                "reason_code",
                "workflow_category",
            )
        }

    def list_resources(
        self,
        kind: ResourceKind,
        *,
        target: str | None = None,
        locale: str | None = None,
    ) -> dict[str, Any]:
        try:
            if kind == "formats":
                projection = runtime_capability_projection(self._controller)
                resources = [
                    {
                        "id": str(source.get("id", "")),
                        "category": str(source.get("category", "")),
                        "available": bool(source.get("available")),
                        "routes": [
                            {
                                "id": str(route.get("id", "")),
                                "source": str(route.get("source", "")),
                                "target": str(route.get("target", "")),
                                "operation": str(route.get("operation", "")),
                                "action": route.get("action") if isinstance(route.get("action"), str) else None,
                                "available": bool(route.get("available")),
                                "state": str(route.get("state", "")),
                                "options": [str(item) for item in route.get("options", [])],
                            }
                            for route in source.get("routes", [])
                            if isinstance(route, dict)
                        ],
                    }
                    for source in projection.get("sources", [])
                    if isinstance(source, dict)
                ]
                return {"kind": kind, "resources": resources}
            if kind == "optimizations":
                projection = runtime_optimization_projection(self._controller)
                raw_resources = projection.get("resources", [])
                if not isinstance(raw_resources, list):
                    raise MachineQueryError("resource_invalid", "Optimization discovery returned invalid data.")
                resources = [
                    {
                        "id": str(item.get("id", "")),
                        "name": str(item.get("name", "")),
                        "description": str(item.get("description", "")),
                        "scopes": [str(scope) for scope in item.get("scopes", [])],
                    }
                    for item in raw_resources
                    if isinstance(item, dict)
                ]
                return {"kind": kind, "resources": resources}
            args = argparse.Namespace(target=target, lang=locale)
            resources = _template_entries(args) if kind == "templates" else _numbering_entries(args, self._controller)
        except CapabilityUnavailableError as exc:
            raise MachineQueryError("resource_unavailable", str(exc)) from exc
        normalized = [
            {
                "id": str(item.get("id", item.get("name", ""))),
                "name": str(item.get("name", item.get("id", ""))),
                "description": str(item.get("description", "")),
                **({"target": str(item["target"])} if item.get("target") else {}),
            }
            for item in resources
        ]
        return {"kind": kind, "resources": normalized}

    def gui_control(
        self,
        action: str,
        *,
        file_path: str | None = None,
        timeout_seconds: int,
    ) -> dict[str, Any]:
        control = self._gui_control
        if control is None:
            raise MachineQueryError("gui_unavailable", "GUI control is unavailable in this assembly.")
        if action not in {"status", "activate", "open"}:
            raise MachineQueryError("gui_action_invalid", "The GUI action is unsupported.")
        try:
            if action == "status":
                raw = control.status(timeout=float(timeout_seconds))
                return {"action": action, "accepted": True, "running": bool(raw.get("running"))}
            if action == "activate":
                control.activate(timeout=float(timeout_seconds))
                return {"action": action, "accepted": True}
            if file_path is not None:
                path = Path(file_path)
                if not path.is_absolute():
                    raise MachineQueryError("invalid_path", "GUI open requires an absolute file path.")
                file_path = str(path.resolve())
            control.open(file_path, timeout=float(timeout_seconds))
            return {"action": action, "accepted": True}
        except GuiControlError as exc:
            raise MachineQueryError(exc.code, str(exc)) from exc

    @staticmethod
    def _verified_file(input_handle: dict[str, Any]) -> Path:
        path = Path(str(input_handle["locator"]["path"]))
        if not path.is_absolute() or not path.is_file():
            raise MachineQueryError("file_not_found", "The input must be an existing absolute file path.")
        digest = hashlib.sha256()
        size_bytes = 0
        try:
            with path.open("rb") as stream:
                while chunk := stream.read(_HASH_CHUNK_BYTES):
                    size_bytes += len(chunk)
                    digest.update(chunk)
        except OSError as exc:
            raise MachineQueryError("file_unreadable", "The input file cannot be read.") from exc
        if size_bytes != int(input_handle["size_bytes"]) or digest.hexdigest() != str(input_handle["sha256"]):
            raise MachineQueryError("input_integrity_mismatch", "The declared input integrity does not match.")
        return path


__all__ = ["MachineQueryError", "MachineQueryService"]
