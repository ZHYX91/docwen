"""Non-semantic metadata added by Office while saving existing content controls."""

from __future__ import annotations

import re
from typing import Any

_WORD = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_ON_OFF = {"true", "false", "1", "0", "on", "off"}
_CONTROL_RUN_PROPERTIES = frozenset(
    f"{_WORD}{name}"
    for name in [
        "b",
        "bCs",
        "bdr",
        "caps",
        "color",
        "cs",
        "dstrike",
        "eastAsianLayout",
        "effect",
        "em",
        "emboss",
        "fitText",
        "highlight",
        "i",
        "iCs",
        "imprint",
        "kern",
        "lang",
        "noProof",
        "oMath",
        "outline",
        "position",
        "rFonts",
        "rStyle",
        "rtl",
        "shadow",
        "shd",
        "smallCaps",
        "snapToGrid",
        "spacing",
        "specVanish",
        "strike",
        "sz",
        "szCs",
        "u",
        "vanish",
        "vertAlign",
        "w",
        "webHidden",
    ]
)


def _control_run_defaults(properties: Any) -> bool:
    # SDT defaults format replacement text, not existing sdtContent. Do not apply
    # these defaults to payload runs or use them as evidence of ownership.
    if properties.attrib or properties.text is not None or properties.tail is not None:
        return False
    seen = set()
    for child in properties:
        if (
            child.tag not in _CONTROL_RUN_PROPERTIES
            or child.tag in seen
            or len(child)
            or child.text is not None
            or child.tail is not None
            or any(not name.startswith(_WORD) for name in child.attrib)
        ):
            return False
        seen.add(child.tag)
    return True


def has_owned_tag_properties(properties: Any, tag: str) -> bool:
    """Prove tag/ID identity while tolerating host control formatting defaults."""
    if properties.attrib or properties.text is not None or properties.tail is not None:
        return False
    tags = properties.findall(f"{_WORD}tag")
    ids = properties.findall(f"{_WORD}id")
    defaults = properties.findall(f"{_WORD}rPr")
    if (
        len(tags) != 1
        or len(ids) > 1
        or len(defaults) > 1
        or len(properties) != len(tags) + len(ids) + len(defaults)
        or any(not _control_run_defaults(item) for item in defaults)
    ):
        return False
    for child in tags + ids:
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
