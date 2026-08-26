"""Test that docwen_application can be imported and has correct dependency direction."""

from unittest.mock import MagicMock

import pytest

pytestmark = pytest.mark.unit


def test_application_importable() -> None:
    """docwen_application should be importable."""
    import docwen_application

    assert docwen_application.__version__ == "0.9.0"


def test_application_submodules_importable() -> None:
    """All application submodules should be importable."""
    from docwen_application import (
        commands,
        controller,
        events,
        optimization_catalog,
        ports,
        runtime_capability_catalog,
        workflows,
    )

    for mod in (
        commands,
        events,
        workflows,
        ports,
        controller,
        optimization_catalog,
        runtime_capability_catalog,
    ):
        assert hasattr(mod, "__name__")


def test_preconversion_package_has_no_public_bypass_api() -> None:
    """Pre-conversion remains an ApplicationController implementation detail."""
    from docwen_application import preconversion

    assert preconversion.__all__ == []


def test_application_depends_on_core() -> None:
    """application should be able to import from docwen_core."""

    # Verify ApplicationController can use core types
    from docwen_application.controller import ApplicationController

    ctrl = ApplicationController()
    assert not ctrl.is_running
    ctrl.start()
    assert ctrl.is_running
    ctrl.stop()
    assert not ctrl.is_running


def test_ports_export_protocols() -> None:
    """HIGH-1 fix: application/ports/ must export RuntimePort, ConfigPort, PresenterPort."""
    from docwen_application.ports import ConfigPort, PresenterPort, RuntimePort

    # All three should be Protocol classes
    for proto in (RuntimePort, ConfigPort, PresenterPort):
        assert hasattr(proto, "__name__")

    # Verify they are Protocols (not concrete classes)
    from typing import Protocol

    assert issubclass(RuntimePort, Protocol)
    assert issubclass(ConfigPort, Protocol)
    assert issubclass(PresenterPort, Protocol)


def test_controller_accepts_dependency_injection() -> None:
    """HIGH-2 fix: ApplicationController must accept port injection at construction."""
    from docwen_application.controller import ApplicationController
    from docwen_application.ports.runtime import ConfigPort, PresenterPort, RuntimePort

    # A controller can exist without a runtime; capability-dependent calls fail closed.
    ctrl = ApplicationController()
    assert not ctrl.has_runtime
    ctrl.start()
    ctrl.stop()

    # 2. Dependencies can be injected at construction time
    mock_runtime = MagicMock(spec=RuntimePort)
    mock_config = MagicMock(spec=ConfigPort)
    mock_presenter = MagicMock(spec=PresenterPort)

    ctrl2 = ApplicationController(
        runtime_port=mock_runtime,
        config_port=mock_config,
        presenter_port=mock_presenter,
    )
    assert ctrl2.has_runtime
    assert ctrl2.config_port is mock_config
    assert ctrl2.presenter_port is mock_presenter
