"""Bundle-owned adapter for CLI-to-GUI runtime/control operations."""

from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import docwen_cli.gui_control_port as gui_control_contract
from docwen_cli.gui_control_port import GuiControlError, GuiControlPort
from docwen_runtime.control import ControlClient, ControlError, ControlNotRunningError, ControlTimeoutError


class GuiControlAdapter(GuiControlPort):
    """Connect the stable CLI contract to the local GUI control endpoint."""

    def __init__(self) -> None:
        self._client = ControlClient(app_name="docwen")

    def status(self, *, timeout: float) -> dict[str, Any]:
        deadline = time.monotonic() + timeout
        try:
            data = self._client.request(
                "status",
                {"_deadline_monotonic": deadline},
                timeout=timeout,
            )
        except ControlNotRunningError:
            return {
                "state": "stopped",
                "running": False,
                "control_ready": False,
                "available": True,
            }
        except ControlError as exc:
            raise _to_cli_error(exc) from exc
        return self._require_running_status(data)

    def activate(self, *, timeout: float) -> dict[str, Any]:
        deadline = time.monotonic() + timeout
        try:
            data = self._client.request(
                "activate",
                {"_deadline_monotonic": deadline},
                timeout=timeout,
            )
        except ControlError as exc:
            raise _to_cli_error(exc) from exc
        return self._require_action_response(data, action="activate")

    def open(self, file_path: str | None, *, timeout: float) -> dict[str, Any]:
        try:
            return self._open(file_path, timeout=timeout)
        except GuiControlError:
            raise
        except ControlError as exc:
            raise _to_cli_error(exc) from exc

    def open_settings(self, section: str, *, timeout: float) -> dict[str, Any]:
        if section not in gui_control_contract.GUI_SETTINGS_SECTIONS:
            raise GuiControlError(
                "settings_section_unavailable",
                "The requested GUI settings section is unavailable.",
                details={
                    "section": section,
                    "available_sections": list(gui_control_contract.GUI_SETTINGS_SECTIONS),
                },
            )
        deadline = time.monotonic() + timeout
        try:
            try:
                status_timeout = max(0.05, timeout)
                status = self._client.request(
                    "status",
                    {"_deadline_monotonic": min(deadline, time.monotonic() + status_timeout)},
                    timeout=status_timeout,
                )
            except ControlNotRunningError:
                self._launch_gui()
                status = self._wait_until_ready(deadline)
            else:
                status = self._require_running_status(status)
            self._require_settings_capability(status, section=section)
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise ControlTimeoutError("Timed out before the GUI could accept the settings request.")
            data = self._client.request(
                "open_settings",
                {
                    "section": section,
                    "_deadline_monotonic": deadline,
                },
                timeout=remaining,
            )
            return self._require_action_response(
                data,
                action="open_settings",
                section=section,
                require_reused=True,
            )
        except GuiControlError:
            raise
        except ControlError as exc:
            raise _to_cli_error(exc) from exc

    def _open(self, file_path: str | None, *, timeout: float) -> dict[str, Any]:
        resolved: Path | None = None
        if file_path is not None:
            path = Path(file_path)
            if not path.is_absolute():
                raise ControlError("invalid_path", "GUI open requires an absolute file path.")
            try:
                resolved = path.resolve(strict=True)
            except OSError as exc:
                raise GuiControlError("file_not_found", "The requested GUI file does not exist.") from exc
            if not resolved.is_file():
                raise GuiControlError("invalid_path", "GUI open requires a regular file.")

        deadline = time.monotonic() + timeout
        try:
            return self._request_open_or_activate(resolved, deadline=deadline)
        except ControlNotRunningError:
            self._launch_gui()

        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise ControlTimeoutError("Timed out while waiting for DocWen GUI control to become ready.")
            try:
                attempt_deadline = min(deadline, time.monotonic() + 0.5)
                status_timeout = attempt_deadline - time.monotonic()
                if status_timeout <= 0:
                    raise ControlTimeoutError("Timed out while waiting for DocWen GUI control to become ready.")
                status = self._client.request(
                    "status",
                    {"_deadline_monotonic": attempt_deadline},
                    timeout=status_timeout,
                )
                if status.get("running") is True and status.get("control_ready") is True:
                    break
            except (ControlNotRunningError, ControlTimeoutError) as exc:
                sleep_remaining = deadline - time.monotonic()
                if sleep_remaining <= 0:
                    raise ControlTimeoutError(
                        "Timed out while waiting for DocWen GUI control to become ready."
                    ) from exc
                time.sleep(min(0.05, sleep_remaining))

        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise ControlTimeoutError("Timed out before the GUI could accept the request.")
        return self._request_open_or_activate(resolved, deadline=deadline)

    def _request_open_or_activate(self, resolved: Path | None, *, deadline: float) -> dict[str, Any]:
        timeout = deadline - time.monotonic()
        if timeout <= 0:
            raise ControlTimeoutError("Timed out before the GUI could accept the request.")
        if resolved is None:
            data = self._client.request(
                "activate",
                {"_deadline_monotonic": deadline},
                timeout=timeout,
            )
            return self._require_action_response(data, action="activate")
        expected_file = str(resolved)
        data = self._client.request(
            "open",
            {
                "file": expected_file,
                "_deadline_monotonic": deadline,
            },
            timeout=timeout,
        )
        return self._require_action_response(data, action="open", expected_file=expected_file)

    def _wait_until_ready(self, deadline: float) -> dict[str, Any]:
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise ControlTimeoutError("Timed out while waiting for DocWen GUI control to become ready.")
            try:
                attempt_deadline = min(deadline, time.monotonic() + 0.5)
                status_timeout = attempt_deadline - time.monotonic()
                if status_timeout <= 0:
                    raise ControlTimeoutError("Timed out while waiting for DocWen GUI control to become ready.")
                status = self._client.request(
                    "status",
                    {"_deadline_monotonic": attempt_deadline},
                    timeout=status_timeout,
                )
                if status.get("running") is True and status.get("control_ready") is True:
                    return self._require_running_status(status)
            except (ControlNotRunningError, ControlTimeoutError) as exc:
                sleep_remaining = deadline - time.monotonic()
                if sleep_remaining <= 0:
                    raise ControlTimeoutError(
                        "Timed out while waiting for DocWen GUI control to become ready."
                    ) from exc
                time.sleep(min(0.05, sleep_remaining))

    @staticmethod
    def _require_settings_capability(status: dict[str, Any], *, section: str) -> None:
        raw_actions = status.get("supported_actions")
        actions = [item for item in raw_actions if isinstance(item, str)] if isinstance(raw_actions, list) else []
        if "open_settings" not in actions:
            raise GuiControlError(
                "capability_unavailable",
                "The running DocWen GUI does not support opening settings. Restart it after upgrading DocWen.",
                details={
                    "required_action": "open_settings",
                    "supported_actions": actions,
                    "restart_required": True,
                    "running": bool(status.get("running")),
                },
            )
        raw_sections = status.get("settings_sections")
        sections = [item for item in raw_sections if isinstance(item, str)] if isinstance(raw_sections, list) else []
        if section not in sections:
            raise GuiControlError(
                "settings_section_unavailable",
                "The running DocWen GUI does not expose the requested settings section.",
                details={"section": section, "available_sections": sections},
            )

    @staticmethod
    def _require_running_status(data: dict[str, Any]) -> dict[str, Any]:
        if data.get("running") is not True or data.get("control_ready") is not True:
            raise _invalid_control_response("status", data)
        return data

    @staticmethod
    def _require_action_response(
        data: dict[str, Any],
        *,
        action: str,
        section: str | None = None,
        expected_file: str | None = None,
        require_reused: bool = False,
    ) -> dict[str, Any]:
        valid = data.get("accepted") is True and data.get("running") is True and data.get("action") == action
        if section is not None:
            valid = valid and data.get("section") == section
        if expected_file is not None:
            valid = valid and data.get("file") == expected_file
        if require_reused:
            valid = valid and type(data.get("reused")) is bool
        if not valid:
            raise _invalid_control_response(action, data)
        return data

    @staticmethod
    def _launch_gui() -> None:
        if getattr(sys, "frozen", False):
            executable = Path(sys.executable).resolve().with_name("DocWen.exe" if os.name == "nt" else "DocWen")
            if not executable.is_file():
                raise ControlError(
                    "gui_start_failed",
                    "DocWen GUI executable is not installed beside DocWenCLI.",
                    details={"expected_name": executable.name},
                )
            command = [str(executable)]
        else:
            command = [sys.executable, "-m", "docwen_gui"]

        creationflags = 0
        start_new_session = os.name != "nt"
        if os.name == "nt":
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
        try:
            subprocess.Popen(
                command,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                close_fds=True,
                shell=False,
                creationflags=creationflags,
                start_new_session=start_new_session,
            )
        except OSError as exc:
            raise ControlError("gui_start_failed", "Unable to start the DocWen GUI.") from exc


def create_gui_control_adapter() -> GuiControlAdapter:
    return GuiControlAdapter()


def _to_cli_error(error: ControlError) -> GuiControlError:
    return GuiControlError(error.code, str(error), details=error.details)


def _invalid_control_response(action: str, data: dict[str, Any]) -> GuiControlError:
    return GuiControlError(
        "invalid_control_response",
        "The running DocWen GUI returned an invalid control response.",
        details={"action": action, "received_fields": sorted(str(key) for key in data)},
    )


__all__ = ["GuiControlAdapter", "create_gui_control_adapter"]
