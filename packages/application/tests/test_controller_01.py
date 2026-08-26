"""Focused tests split from test_controller.py."""

from __future__ import annotations

import pytest

from ._controller_support import (
    ApplicationController,
    MagicMock,
)
from ._controller_support import (
    _isolate_controller_tests_from_core_admission as _isolate_controller_tests_from_core_admission,
)
from ._controller_support import (
    mock_config as mock_config,
)
from ._controller_support import (
    mock_presenter as mock_presenter,
)
from ._controller_support import (
    mock_runtime as mock_runtime,
)

pytestmark = pytest.mark.unit


class TestControllerLifecycle:
    """start/stop/is_running/has_runtime."""

    def test_initial_state_is_stopped(self) -> None:
        ctrl = ApplicationController()
        assert not ctrl.is_running

    def test_start_sets_running(self) -> None:
        ctrl = ApplicationController()
        ctrl.start()
        assert ctrl.is_running

    def test_stop_clears_running(self) -> None:
        ctrl = ApplicationController()
        ctrl.start()
        ctrl.stop()
        assert not ctrl.is_running

    def test_start_stop_idempotent(self) -> None:
        ctrl = ApplicationController()
        ctrl.start()
        ctrl.start()
        assert ctrl.is_running
        ctrl.stop()
        ctrl.stop()
        assert not ctrl.is_running

    def test_has_runtime_false_by_default(self) -> None:
        ctrl = ApplicationController()
        assert not ctrl.has_runtime

    def test_has_runtime_true_when_injected(self, mock_runtime: MagicMock) -> None:
        ctrl = ApplicationController(runtime_port=mock_runtime)
        assert ctrl.has_runtime


class TestDependencyInjection:
    """Port injection at construction time."""

    def test_no_arg_construction(self) -> None:
        ctrl = ApplicationController()
        assert not ctrl.has_runtime
        assert ctrl.config_port is None
        assert ctrl.presenter_port is None

    def test_partial_injection_runtime_only(self, mock_runtime: MagicMock) -> None:
        ctrl = ApplicationController(runtime_port=mock_runtime)
        assert ctrl.has_runtime
        assert ctrl.config_port is None
        assert ctrl.presenter_port is None

    def test_full_injection(
        self,
        mock_runtime: MagicMock,
        mock_config: MagicMock,
        mock_presenter: MagicMock,
    ) -> None:
        ctrl = ApplicationController(
            runtime_port=mock_runtime,
            config_port=mock_config,
            presenter_port=mock_presenter,
        )
        assert ctrl.has_runtime
        assert ctrl.config_port is mock_config
        assert ctrl.presenter_port is mock_presenter

    def test_stop_releases_injected_runtime(self, mock_runtime: MagicMock) -> None:
        ctrl = ApplicationController(runtime_port=mock_runtime)
        ctrl.start()

        ctrl.stop()

        assert not ctrl.is_running
        assert not ctrl.has_runtime
        mock_runtime.shutdown.assert_called_once_with()

    def test_injected_ports_referentially_transparent(self, mock_runtime: MagicMock) -> None:
        ctrl1 = ApplicationController(runtime_port=mock_runtime)
        ctrl2 = ApplicationController()
        assert ctrl1.has_runtime
        assert not ctrl2.has_runtime


def test_raw_command_factories_are_not_public_execution_bypasses() -> None:
    ctrl = ApplicationController()

    assert not hasattr(ctrl, "convert_command")
    assert not hasattr(ctrl, "batch_command")
    assert not hasattr(ctrl, "aggregate_command")


def test_application_root_exports_only_the_admitted_execution_boundaries() -> None:
    import docwen_application
    import docwen_application.commands
    import docwen_application.workflows

    assert docwen_application.__all__ == [
        "ApplicationController",
        "ControllerError",
        "ConversionService",
        "ConversionServiceError",
    ]
    assert docwen_application.commands.__all__ == []
    assert docwen_application.workflows.__all__ == []
