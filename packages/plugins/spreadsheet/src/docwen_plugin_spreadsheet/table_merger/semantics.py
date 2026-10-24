"""Value identity, coverage and bounded decision evidence for table merging."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any


def value_identity(value: Any, *, source_type: str | None = None) -> Any:
    """Compare identities without interpreting textual IDs as floating numbers.

    Only blank text is empty. Python's numeric equality compares integers and
    floats without first rounding an integer to a float; booleans stay separate.
    A saved Excel error is distinct from ordinary text with the same spelling.
    Arithmetic in cell mode has its own, deliberately different policy.
    """
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if source_type == "e":
        return ("error", str(value))
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)):
        return ("num", value)
    if isinstance(value, (date, datetime)):
        return ("date", value.isoformat())
    return ("str", str(value))


def coverage_relation(*, base_unique: bool, collect_unique: bool) -> str:
    """Classify non-conflicting records after all aligned fields were compared."""
    if base_unique and collect_unique:
        return "complementary"
    if collect_unique:
        return "collect_covers_base"
    if base_unique:
        return "base_covers_collect"
    return "equal"


@dataclass
class MergeSummary:
    """Count decisions with bounded source locations, never cell contents.

    Locations refer to the collection input before insertion; output row and
    column numbers can move again as later records are inserted.
    """

    source_name: str
    sheet: str
    mode: str
    row_offset: int
    col_offset: int
    counts: dict[str, int] = field(
        default_factory=lambda: dict.fromkeys(("added", "replaced", "skipped", "combined"), 0)
    )
    locations: dict[str, list[str]] = field(default_factory=dict)

    def record(self, decision: str, location: str) -> None:
        self.counts[decision] += 1
        samples = self.locations.setdefault(decision, [])
        if len(samples) < 20:
            samples.append(location)

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_name": self.source_name,
            "sheet": self.sheet,
            "mode": self.mode,
            "row_offset": self.row_offset,
            "col_offset": self.col_offset,
            "counts": dict(self.counts),
            "locations": {key: list(value) for key, value in self.locations.items()},
            "unknown": 0,
        }

    def message(self) -> str:
        counts = ", ".join(f"{key}={count}" for key, count in self.counts.items())
        samples = "; ".join(f"{key}: {', '.join(value)}" for key, value in self.locations.items())
        return (
            f"{self.source_name} / {self.sheet} ({self.mode}; offset {self.row_offset}, {self.col_offset}): "
            f"{counts}, unknown=0. Collection locations (up to 20 per decision): {samples}."
        )
