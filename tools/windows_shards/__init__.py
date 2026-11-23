"""File selection for the Windows full-suite jobs."""

from __future__ import annotations

GROUPS = ("control", "apps-plugins", "remaining")


def selected(group: str, nodeid: str) -> bool:
    if group == "control":
        return True
    path = nodeid.split("::", 1)[0].replace("\\", "/")
    if path.startswith(("packages/apps/", "packages/plugins/")):
        return group == "apps-plugins"
    return group == "remaining"
