"""Test that docwen_bundle can be imported."""

import pytest

pytestmark = pytest.mark.unit


def test_bundle_importable() -> None:
    """docwen_bundle should be importable."""
    import docwen_bundle

    assert docwen_bundle.__version__ == "0.9.1"
