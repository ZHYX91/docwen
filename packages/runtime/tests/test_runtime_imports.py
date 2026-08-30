"""Test that docwen_runtime can be imported and has correct submodule structure."""

import pytest

pytestmark = pytest.mark.unit


def test_runtime_importable() -> None:
    """docwen_runtime should be importable."""
    import docwen_runtime

    assert docwen_runtime.__version__ == "0.9.1"


def test_runtime_submodules_importable() -> None:
    """All runtime submodules should be importable."""
    from docwen_runtime import config, engine, output, plugin_registry, workspace
    from docwen_runtime import logging as runtime_logging

    for mod in (engine, plugin_registry, output, workspace, config, runtime_logging):
        assert hasattr(mod, "__name__")
