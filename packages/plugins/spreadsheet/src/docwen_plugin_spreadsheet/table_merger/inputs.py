"""Read merge inputs without treating unavailable formula results as blanks."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen_plugin_spreadsheet.csv_xlsx.converter import _find_unavailable_formula_caches, _load_xlsx_views


class MissingFormulaCacheError(ValueError):
    """No output may be published from this incomplete cached-value view."""

    def __init__(self, source: str, count: int, locations: list[str]) -> None:
        self.location = f"{Path(source).name}: {', '.join(locations)}"
        self.count = count
        omitted = f" ({count - len(locations)} more not listed)" if count > len(locations) else ""
        super().__init__(
            f"Cannot merge: {count} formula cell(s) have no usable cached value in {self.location}{omitted}. "
            "Recalculate and save the workbook in a spreadsheet application, then retry. "
            "DocWen does not calculate formulas or verify that existing caches are current."
        )


def load_merge_workbook(
    path: str,
    *,
    base: bool,
    cancel_check: Callable[[], None],
    source_name: str | None = None,
) -> Any:
    """Check every retained base sheet, or only the collected active sheet."""
    values, formulas = _load_xlsx_views(path)
    try:
        active = formulas.active
        sheets = None if base else ([active.title] if active is not None else [])
        count, locations = _find_unavailable_formula_caches(
            values, formulas, input_path=path, sheet_names=sheets, cancel_check=cancel_check
        )
        if count:
            raise MissingFormulaCacheError(source_name or path, count, locations)
    except BaseException:
        values.close()
        raise
    finally:
        formulas.close()
    return values
