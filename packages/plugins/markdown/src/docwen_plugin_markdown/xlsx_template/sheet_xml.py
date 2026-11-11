"""Schema ordering shared by worksheet values and object carriers."""

from __future__ import annotations

from typing import Any

from lxml import etree

_WORKSHEET_ORDER = (
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
)


def insert_sheet_child(root: Any, child: Any) -> None:
    """Insert a worksheet child before later schema members and extensions."""
    order = {name: index for index, name in enumerate(_WORKSHEET_ORDER)}
    rank = order.get(etree.QName(child).localname, len(order))
    for index, current in enumerate(root):
        if not isinstance(current.tag, str):
            continue
        if order.get(etree.QName(current).localname, len(order)) > rank:
            root.insert(index, child)
            return
    root.append(child)
