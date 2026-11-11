"""One schema-ordered append operation for generated Word numbering definitions."""

from __future__ import annotations

from collections.abc import Iterable
from copy import deepcopy

from lxml import etree

from docwen_core.docx_package_parts import WML_NS

_ORDER = {
    f"{{{WML_NS}}}{name}": rank for rank, name in enumerate(("numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"))
}


def validate_numbering_order(root: etree._Element) -> None:
    """Check known CT_Numbering children without interpreting opaque extensions."""
    if root.tag != f"{{{WML_NS}}}numbering":
        raise ValueError("Invalid numbering root")
    order = [_ORDER[item.tag] for item in root if item.tag in _ORDER]
    if order != sorted(order) or order.count(3) > 1:
        raise ValueError("Numbering children violate CT_Numbering order")


def append_numbering_definitions(
    root: etree._Element,
    abstracts: Iterable[etree._Element],
    instances: Iterable[etree._Element],
) -> None:
    """Preserve existing nodes and append copies before the trailing cleanup node."""
    validate_numbering_order(root)
    for source in abstracts:
        if source.tag != f"{{{WML_NS}}}abstractNum":
            raise ValueError("Expected a generated abstractNum definition")
        following = next((item for item in root if _ORDER.get(item.tag, -1) >= 2), None)
        element = deepcopy(source)
        if following is None:
            root.append(element)
        else:
            following.addprevious(element)
    for source in instances:
        if source.tag != f"{{{WML_NS}}}num":
            raise ValueError("Expected a generated num instance")
        cleanup = root.find(f"{{{WML_NS}}}numIdMacAtCleanup")
        element = deepcopy(source)
        if cleanup is None:
            root.append(element)
        else:
            cleanup.addprevious(element)
    validate_numbering_order(root)
