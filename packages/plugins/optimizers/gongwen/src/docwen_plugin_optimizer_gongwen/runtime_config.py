"""Runtime configuration for gongwen content generation."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class GongwenContentRuntimeConfig:
    """Configuration for Gongwen Markdown content generation."""

    horizontal_rule_enabled: bool = False
    page_break_marker: str = "---"
    horizontal_rule_marker: str = "---"
    section_break_marker: str = "---"


# Default config instance
DEFAULT_CONFIG = GongwenContentRuntimeConfig()


def configure_gongwen_content_runtime(**kwargs) -> GongwenContentRuntimeConfig:
    """Create a GongwenContentRuntimeConfig with custom settings."""
    return GongwenContentRuntimeConfig(**kwargs)
