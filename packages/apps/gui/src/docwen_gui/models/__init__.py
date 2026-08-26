"""Typed config models for the DocWen GUI Settings.

These dataclasses are the typed configuration model that Settings UI
bidirectionally binds to.  They serve as the state source of truth
within the SettingsViewModel and are independent of any specific
config backend (TOML, Pydantic, etc.).

The ApplicationController (or a mock in tests) provides the initial
values and receives finalized values when the user applies settings.
"""

from .settings_config import (
    ConversionDefaultsConfig,
    ExportConfig,
    FormattingConfig,
    GUIConfig,
    LinkConfig,
    LoggingConfig,
    OutputConfig,
    ProofreadConfig,
    SettingsConfig,
    SoftwarePriorityConfig,
    TextConfig,
)

__all__ = [
    "ConversionDefaultsConfig",
    "ExportConfig",
    "FormattingConfig",
    "GUIConfig",
    "LinkConfig",
    "LoggingConfig",
    "OutputConfig",
    "ProofreadConfig",
    "SettingsConfig",
    "SoftwarePriorityConfig",
    "TextConfig",
]
