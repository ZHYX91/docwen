"""Internal application-owned pre-conversion implementation.

Callers execute admitted requests through :class:`ApplicationController`.
This package deliberately exports no path-plus-format shortcut that could
bypass the controller's content-derived file identity and admission policy.
"""

__all__: list[str] = []
