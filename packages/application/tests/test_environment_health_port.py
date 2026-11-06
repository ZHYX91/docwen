"""Controller delegates runtime health checks only via the optional typed port."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from docwen_application.controller import ApplicationController

pytestmark = pytest.mark.unit


def test_no_runtime_health_port_is_explicit_unavailable() -> None:
    controller = ApplicationController(runtime_port=SimpleNamespace())
    assert controller.probe_runtime_workspace()["reason"] == "probe_unavailable"
    assert controller.probe_output_directory("/test")["reason"] == "probe_unavailable"


def test_controller_forwards_runtime_probes_without_exposing_implementation() -> None:
    called = []

    class Port:
        def probe_runtime_workspace(self):
            called.append("runtime")
            return {"status": "ok", "reason": None}

        def probe_output_directory(self, target):
            called.append(target)
            return {"status": "fail", "reason": "not_writable"}

    controller = ApplicationController(runtime_port=Port())
    assert controller.probe_runtime_workspace()["status"] == "ok"
    assert controller.probe_output_directory("/check")["reason"] == "not_writable"
    assert called == ["runtime", "/check"]
