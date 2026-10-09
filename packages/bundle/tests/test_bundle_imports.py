"""Test that docwen_bundle can be imported."""

import pytest

from docwen_core.version import PRODUCT_VERSION

pytestmark = pytest.mark.unit


def test_bundle_importable() -> None:
    """docwen_bundle should be importable."""
    import docwen_bundle

    assert docwen_bundle.__version__ == PRODUCT_VERSION
