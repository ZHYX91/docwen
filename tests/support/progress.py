"""Fake implementation of ProgressSink protocol."""

from __future__ import annotations


class FakeProgressSink:
    """Fake progress sink — records progress and artifact events for assertions."""

    def __init__(self) -> None:
        self.events: list[tuple[float, str]] = []
        self.artifacts: list[tuple[str, str]] = []
        self.diagnostics: list[tuple[str, str, str, str]] = []

    def report_progress(self, percent: float, message: str = "") -> None:
        self.events.append((percent, message))

    def report_diagnostic(self, level: str, message: str, code: str = "", location: str = "") -> None:
        self.diagnostics.append((level, message, code, location))

    def report_artifact_ready(self, artifact_id: str, suggested_name: str) -> None:
        self.artifacts.append((artifact_id, suggested_name))
