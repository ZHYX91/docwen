"""DocWen Runtime package.

Package import stays deliberately lightweight so the bundle composition root
can install the dependency egress guard before loading runtime services.
Submodules remain importable through normal ``from docwen_runtime import x``
semantics; the version is resolved lazily on first access.
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
