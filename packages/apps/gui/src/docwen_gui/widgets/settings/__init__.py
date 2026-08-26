"""Settings dialog widgets.

Each tab is a self-contained QWidget that binds to ``SettingsViewModel``.
The ``SettingsDialog`` owns the tabs and coordinates Apply/Cancel/Reset.
"""

from .dialog import SettingsDialog

__all__ = [
    "SettingsDialog",
]
