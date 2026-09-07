"""Theme-independent control geometry shared by every component stylesheet.

Themes supply colours. This module owns the complete QSS box model so that
reapplying a theme cannot lose padding or change the minimum outer dimensions.
"""

from __future__ import annotations

from .design_tokens import Border, Radius, Sizing, Spacing


def button_geometry_qss(*, minimum_height: int = Sizing.CONTROL_HEIGHT, minimum_width: int | None = None) -> str:
    """Translate outer widget dimensions into one complete QSS button box."""
    content_height = minimum_height - 2 * (Spacing.XS + Border.THIN)
    declarations = [
        f"    border-radius: {Radius.MEDIUM}px;",
        f"    padding: {Spacing.XS}px {Spacing.MD}px;",
        f"    min-height: {content_height}px;",
    ]
    if minimum_width is not None:
        content_width = minimum_width - 2 * (Spacing.MD + Border.THIN)
        declarations.append(f"    min-width: {content_width}px;")
    return "\n".join(declarations)
