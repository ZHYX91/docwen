"""Pure proofread rule data injected by the runtime."""

from __future__ import annotations

from dataclasses import dataclass, field


def _copy_map(values: dict[str, tuple[str, ...]] | None = None) -> dict[str, tuple[str, ...]]:
    if not values:
        return {}
    return {str(key): tuple(str(item) for item in items) for key, items in values.items()}


@dataclass(slots=True)
class ProofreadRules:
    """Read-only-ish proofread rule bundle passed into plugins as pure data."""

    symbol_pairs: tuple[tuple[str, str], ...] = ()
    symbol_map: dict[str, tuple[str, ...]] = field(default_factory=dict)
    typos_map: dict[str, tuple[str, ...]] = field(default_factory=dict)
    sensitive_words: dict[str, tuple[str, ...]] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.symbol_pairs = tuple((str(opening), str(closing)) for opening, closing in self.symbol_pairs)
        self.symbol_map = _copy_map(self.symbol_map)
        self.typos_map = _copy_map(self.typos_map)
        self.sensitive_words = _copy_map(self.sensitive_words)
