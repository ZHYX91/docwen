"""Non-semantic metadata added by Office while saving existing content controls."""

from __future__ import annotations

import re
from typing import Any

_WORD = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_ON_OFF = {"true", "false", "1", "0", "on", "off"}


def has_owned_tag_properties(properties: Any, tag: str) -> bool:
    """Require the exact authored tag and at most one valid host-assigned SDT ID."""
    if properties.attrib or properties.text is not None or properties.tail is not None:
        return False
    tags = properties.findall(f"{_WORD}tag")
    ids = properties.findall(f"{_WORD}id")
    if len(tags) != 1 or len(ids) > 1 or len(properties) != len(tags) + len(ids):
        return False
    for child in properties:
        if set(child.attrib) != {f"{_WORD}val"} or child.text is not None or child.tail is not None or len(child):
            return False
    if tags[0].get(f"{_WORD}val") != tag:
        return False
    if ids:
        value = ids[0].get(f"{_WORD}val", "")
        if re.fullmatch(r"-?(?:0|[1-9][0-9]{0,9})", value) is None or not -(2**31) <= int(value) < 2**31:
            return False
    return True


def field_run_payload(run: Any) -> list[Any] | None:
    """Exclude valid revision IDs and the proofing hint Word adds to field caches."""
    allowed = {f"{_WORD}rsidR", f"{_WORD}rsidRPr", f"{_WORD}rsidDel"}
    if not set(run.attrib).issubset(allowed) or any(
        re.fullmatch(r"[0-9A-Fa-f]{8}", value) is None for value in run.attrib.values()
    ):
        return None
    if run.xpath("text()") or run.tail is not None:
        return None
    children = list(run)
    if children and children[0].tag == f"{_WORD}rPr":
        properties = children.pop(0)
        if properties.attrib or properties.text is not None or properties.tail is not None or len(properties) > 1:
            return None
        for item in properties:
            if (
                item.tag != f"{_WORD}noProof"
                or not set(item.attrib).issubset({f"{_WORD}val"})
                or item.get(f"{_WORD}val", "true") not in _ON_OFF
                or item.text is not None
                or item.tail is not None
                or len(item)
            ):
                return None
    return children
