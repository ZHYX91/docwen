"""Evidence for the optional Windows full-suite partition experiment."""

from __future__ import annotations

GROUPS = ("control", "apps-plugins", "remaining")


def selected(group: str, nodeid: str) -> bool:
    path = nodeid.split("::", 1)[0].replace("\\", "/")
    in_apps_plugins = path.startswith(("packages/apps/", "packages/plugins/"))
    return group == "control" or (group == "apps-plugins") == in_apps_plugins
