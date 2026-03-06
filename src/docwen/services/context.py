from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol


class Translator(Protocol):
    def __call__(self, key: str, default: str | None = None, **kwargs: Any) -> str: ...

class ActualFormatDetector(Protocol):
    def __call__(self, file_path: str) -> str: ...


class CategoryDetector(Protocol):
    def __call__(self, file_path: str, actual_format: str | None = None) -> str: ...


class StrategyGetter(Protocol):
    def __call__(
        self,
        *,
        action_type: str | None,
        source_format: str | None,
        target_format: str | None,
    ) -> Any: ...


@dataclass(frozen=True)
class AppContext:
    t: Translator
    config_manager: Any
    detect_actual_file_format: ActualFormatDetector
    get_actual_file_category: CategoryDetector
    get_strategy: StrategyGetter


def get_default_context() -> AppContext:
    from docwen.config.config_manager import config_manager
    from docwen.services.strategies import get_strategy
    from docwen.translation import t
    from docwen.utils.file_type_utils import detect_actual_file_format, get_strategy_file_category

    return AppContext(
        t=t,
        config_manager=config_manager,
        detect_actual_file_format=detect_actual_file_format,
        get_actual_file_category=get_strategy_file_category,
        get_strategy=get_strategy,
    )
