"""DocWen product version.

This module is the single runtime source for the version displayed by the
desktop application, CLI, and assembled bundle.  Package metadata is checked
against this value by repository contract tests.
"""

from __future__ import annotations

PRODUCT_VERSION = "0.9.1"
__version__ = PRODUCT_VERSION

__all__ = ["PRODUCT_VERSION", "__version__"]
