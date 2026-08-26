"""Test that docwen_core can be imported and its public API is available."""

import pytest

pytestmark = pytest.mark.unit


def test_core_importable() -> None:
    """docwen_core should be importable."""
    import docwen_core

    assert docwen_core.__version__ == "0.9.0"


def test_core_submodules_importable() -> None:
    """All core submodules should be importable."""
    from docwen_core import cancellation, errors, events, formats, models, options, protocols

    # Verify they are modules
    for mod in (models, protocols, formats, events, options, errors, cancellation):
        assert hasattr(mod, "__name__")


def test_core_errors_hierarchy() -> None:
    """Core error classes should have the correct hierarchy."""
    from docwen_core.errors import (
        CancellationRequested,
        ConfigurationError,
        ConversionError,
        DocWenError,
        ValidationError,
    )

    assert issubclass(ConversionError, DocWenError)
    assert issubclass(ConfigurationError, DocWenError)
    assert issubclass(ValidationError, DocWenError)
    assert issubclass(CancellationRequested, DocWenError)
