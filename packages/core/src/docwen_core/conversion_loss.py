"""Request-local, output-neutral records of unintended conversion loss."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, replace

from docwen_core.models.result import ConversionDiagnostic


def loss_text(messages: Mapping[str, str] | None, key: str, fallback: str) -> str:
    """Resolve request-owned output copy without importing runtime localization."""
    return (messages or {}).get(f"conversion_loss.{key}", fallback)


@dataclass(frozen=True, slots=True)
class ConversionLoss:
    """One actual loss, with readable source and output coordinates."""

    code: str
    reason: str
    change: str
    source_location: str = ""
    output_location: str = ""
    original: str = ""

    def message(self, messages: Mapping[str, str] | None = None) -> str:
        original = (
            loss_text(messages, "original", "Original: {original}").format(original=self.original)
            if self.original
            else ""
        )
        location = "; ".join(dict.fromkeys(filter(None, (self.source_location, self.output_location))))
        return (
            loss_text(messages, "comment", "{reason}\n{change}\n{location}\n{original}")
            .format(reason=self.reason, change=self.change, location=location, original=original)
            .strip()
        )

    def to_diagnostic(
        self, messages: Mapping[str, str] | None = None, *, include_source_content: bool = True
    ) -> ConversionDiagnostic:
        loss = self if include_source_content else replace(self, original="", source_location="", output_location="")
        return ConversionDiagnostic(
            level="warning",
            code=self.code,
            message=loss.message(messages),
            location=loss.output_location or loss.source_location,
        )


class LossCollector:
    """Deduplicate identical records, preserving occurrence order."""

    def __init__(self) -> None:
        self._records: dict[ConversionLoss, None] = {}

    def add(self, loss: ConversionLoss) -> bool:
        if loss in self._records:
            return False
        self._records[loss] = None
        return True

    @property
    def records(self) -> tuple[ConversionLoss, ...]:
        return tuple(self._records)

    @property
    def diagnostics(self) -> tuple[ConversionDiagnostic, ...]:
        return tuple(loss.to_diagnostic() for loss in self._records)

    def summary(
        self, messages: Mapping[str, str] | None = None, *, annotated_count: int = 0
    ) -> ConversionDiagnostic | None:
        count = len(self._records)
        if not count:
            return None
        if annotated_count == count:
            message = loss_text(
                messages,
                "summary_annotated",
                "Conversion completed with {count} items not fully preserved. Comments were added.",
            )
        else:
            message = loss_text(messages, "summary", "Conversion completed with {count} items not fully preserved.")
        return ConversionDiagnostic(
            level="warning",
            code="CONVERSION-LOSS-SUMMARY",
            message=message.format(count=count),
        )
