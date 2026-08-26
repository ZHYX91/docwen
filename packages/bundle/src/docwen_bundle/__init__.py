"""DocWen Bundle composition package.

The package initializer must remain stdlib-only: entry modules install the
dependency egress guard before importing the product composition.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_core.version import __version__ as __version__

__all__ = ["__version__"]


def __getattr__(name: str) -> Any:
    if name == "__version__":
        from docwen_core.version import __version__

        return __version__
    raise AttributeError(name)
