"""Non-semantic metadata added by Office while saving existing content controls."""

from __future__ import annotations

import re
from typing import Any

_WORD = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_ON_OFF = {"true", "false", "1", "0", "on", "off"}
_TOGGLE_PROPERTIES = frozenset(
    [
        "b",
        "bCs",
        "caps",
        "cs",
        "dstrike",
        "emboss",
        "i",
        "iCs",
        "imprint",
        "noProof",
        "oMath",
        "outline",
        "rtl",
        "shadow",
        "smallCaps",
        "snapToGrid",
        "specVanish",
        "strike",
        "vanish",
        "webHidden",
    ]
)
_THEME_COLORS = frozenset(
    [
        "dark1",
        "light1",
        "dark2",
        "light2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hyperlink",
        "followedHyperlink",
        "none",
        "background1",
        "text1",
        "background2",
        "text2",
    ]
)
_ENUM_PROPERTIES = {
    "highlight": frozenset(
        [
            "black",
            "blue",
            "cyan",
            "green",
            "magenta",
            "red",
            "yellow",
            "white",
            "darkBlue",
            "darkCyan",
            "darkGreen",
            "darkMagenta",
            "darkRed",
            "darkYellow",
            "darkGray",
            "lightGray",
            "none",
        ]
    ),
    "effect": frozenset(["blinkBackground", "lights", "antsBlack", "antsRed", "shimmer", "sparkle", "none"]),
    "em": frozenset(["none", "dot", "comma", "circle", "underDot"]),
    "vertAlign": frozenset(["baseline", "superscript", "subscript"]),
}
_UNDERLINES = frozenset(
    [
        "single",
        "words",
        "double",
        "thick",
        "dotted",
        "dottedHeavy",
        "dash",
        "dashedHeavy",
        "dashLong",
        "dashLongHeavy",
        "dotDash",
        "dashDotHeavy",
        "dotDotDash",
        "dashDotDotHeavy",
        "wave",
        "wavyHeavy",
        "wavyDouble",
        "none",
    ]
)
_THEME_FONTS = frozenset(
    ["majorEastAsia", "majorBidi", "majorAscii", "majorHAnsi", "minorEastAsia", "minorBidi", "minorAscii", "minorHAnsi"]
)
_SHADING_PATTERNS = frozenset(
    [
        "nil",
        "clear",
        "solid",
        "horzStripe",
        "vertStripe",
        "reverseDiagStripe",
        "diagStripe",
        "horzCross",
        "diagCross",
        "thinHorzStripe",
        "thinVertStripe",
        "thinReverseDiagStripe",
        "thinDiagStripe",
        "thinHorzCross",
        "thinDiagCross",
        "pct5",
        "pct10",
        "pct12",
        "pct15",
        "pct20",
        "pct25",
        "pct30",
        "pct35",
        "pct37",
        "pct40",
        "pct45",
        "pct50",
        "pct55",
        "pct60",
        "pct62",
        "pct65",
        "pct70",
        "pct75",
        "pct80",
        "pct85",
        "pct87",
        "pct90",
        "pct95",
    ]
)
_BORDERS = frozenset(
    [
        "nil",
        "none",
        "single",
        "thick",
        "double",
        "dotted",
        "dashed",
        "dotDash",
        "dotDotDash",
        "triple",
        "thinThickSmallGap",
        "thickThinSmallGap",
        "thinThickThinSmallGap",
        "thinThickMediumGap",
        "thickThinMediumGap",
        "thinThickThinMediumGap",
        "thinThickLargeGap",
        "thickThinLargeGap",
        "thinThickThinLargeGap",
        "wave",
        "doubleWave",
        "dashSmallGap",
        "dashDotStroked",
        "threeDEmboss",
        "threeDEngrave",
        "outset",
        "inset",
        "apples",
        "archedScallops",
        "babyPacifier",
        "babyRattle",
        "balloons3Colors",
        "balloonsHotAir",
        "basicBlackDashes",
        "basicBlackDots",
        "basicBlackSquares",
        "basicThinLines",
        "basicWhiteDashes",
        "basicWhiteDots",
        "basicWhiteSquares",
        "basicWideInline",
        "basicWideMidline",
        "basicWideOutline",
        "bats",
        "birds",
        "birdsFlight",
        "cabins",
        "cakeSlice",
        "candyCorn",
        "celticKnotwork",
        "certificateBanner",
        "chainLink",
        "champagneBottle",
        "checkedBarBlack",
        "checkedBarColor",
        "checkered",
        "christmasTree",
        "circlesLines",
        "circlesRectangles",
        "classicalWave",
        "clocks",
        "compass",
        "confetti",
        "confettiGrays",
        "confettiOutline",
        "confettiStreamers",
        "confettiWhite",
        "cornerTriangles",
        "couponCutoutDashes",
        "couponCutoutDots",
        "crazyMaze",
        "creaturesButterfly",
        "creaturesFish",
        "creaturesInsects",
        "creaturesLadyBug",
        "crossStitch",
        "cup",
        "decoArch",
        "decoArchColor",
        "decoBlocks",
        "diamondsGray",
        "doubleD",
        "doubleDiamonds",
        "earth1",
        "earth2",
        "eclipsingSquares1",
        "eclipsingSquares2",
        "eggsBlack",
        "fans",
        "film",
        "firecrackers",
        "flowersBlockPrint",
        "flowersDaisies",
        "flowersModern1",
        "flowersModern2",
        "flowersPansy",
        "flowersRedRose",
        "flowersRoses",
        "flowersTeacup",
        "flowersTiny",
        "gems",
        "gingerbreadMan",
        "gradient",
        "handmade1",
        "handmade2",
        "heartBalloon",
        "heartGray",
        "hearts",
        "heebieJeebies",
        "holly",
        "houseFunky",
        "hypnotic",
        "iceCreamCones",
        "lightBulb",
        "lightning1",
        "lightning2",
        "mapPins",
        "mapleLeaf",
        "mapleMuffins",
        "marquee",
        "marqueeToothed",
        "moons",
        "mosaic",
        "musicNotes",
        "northwest",
        "ovals",
        "packages",
        "palmsBlack",
        "palmsColor",
        "paperClips",
        "papyrus",
        "partyFavor",
        "partyGlass",
        "pencils",
        "people",
        "peopleWaving",
        "peopleHats",
        "poinsettias",
        "postageStamp",
        "pumpkin1",
        "pushPinNote2",
        "pushPinNote1",
        "pyramids",
        "pyramidsAbove",
        "quadrants",
        "rings",
        "safari",
        "sawtooth",
        "sawtoothGray",
        "scaredCat",
        "seattle",
        "shadowedSquares",
        "sharksTeeth",
        "shorebirdTracks",
        "skyrocket",
        "snowflakeFancy",
        "snowflakes",
        "sombrero",
        "southwest",
        "stars",
        "starsTop",
        "stars3d",
        "starsBlack",
        "starsShadowed",
        "sun",
        "swirligig",
        "tornPaper",
        "tornPaperBlack",
        "trees",
        "triangleParty",
        "triangles",
        "tribal1",
        "tribal2",
        "tribal3",
        "tribal4",
        "tribal5",
        "tribal6",
        "triangle1",
        "triangle2",
        "triangleCircle1",
        "triangleCircle2",
        "shapes1",
        "shapes2",
        "twistedLines1",
        "twistedLines2",
        "vine",
        "waveline",
        "weavingAngles",
        "weavingBraid",
        "weavingRibbon",
        "weavingStrips",
        "whiteFlowers",
        "woodwork",
        "xIllusions",
        "zanyTriangles",
        "zigZag",
        "zigZagStitch",
    ]
)
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


def _control_property_values(child: Any) -> bool:
    name = child.tag.removeprefix(_WORD)
    attributes = {key.removeprefix(_WORD): value for key, value in child.attrib.items()}
    if name in _TOGGLE_PROPERTIES:
        return set(attributes).issubset({"val"}) and attributes.get("val", "true") in _ON_OFF
    if name in {"sz", "szCs", "kern"}:
        return set(attributes) == {"val"} and _unsigned_measure(attributes["val"])
    if name == "color":
        if not set(attributes).issubset({"val", "themeColor", "themeTint", "themeShade"}):
            return False
        if re.fullmatch(r"auto|[0-9A-Fa-f]{6}", attributes.get("val", "")) is None:
            return False
        if "themeColor" in attributes and attributes["themeColor"] not in _THEME_COLORS:
            return False
        return all(
            re.fullmatch(r"[0-9A-Fa-f]{2}", attributes[key]) is not None
            for key in ("themeTint", "themeShade")
            if key in attributes
        )
    if name in _ENUM_PROPERTIES:
        return set(attributes) == {"val"} and attributes["val"] in _ENUM_PROPERTIES[name]
    if name in {"position", "spacing"}:
        return (
            set(attributes) == {"val"}
            and re.fullmatch(r"(?:[+-]?[0-9]+|-?[0-9]+(?:\.[0-9]+)?(?:mm|cm|in|pt|pc|pi))", attributes["val"])
            is not None
        )
    if name == "w":
        value = attributes.get("val", "100").strip(" \t\r\n")
        return set(attributes).issubset({"val"}) and (
            _integer_in_range(value, 0, 600) or re.fullmatch(r"0*(?:600|[0-5]?[0-9]?[0-9])%", value) is not None
        )
    if name == "rStyle":
        return set(attributes) == {"val"}
    if name == "lang":
        return set(attributes).issubset({"val", "eastAsia", "bidi"})
    if name == "rFonts":
        return _font_attributes(attributes)
    if name == "fitText":
        return (
            set(attributes).issubset({"val", "id"})
            and _unsigned_measure(attributes.get("val", ""))
            and ("id" not in attributes or _decimal_integer(attributes["id"]))
        )
    if name == "eastAsianLayout":
        return _east_asian_attributes(attributes)
    if name == "u":
        return (
            set(attributes).issubset({"val", "color", "themeColor", "themeTint", "themeShade"})
            and attributes.get("val", "single") in _UNDERLINES
            and _color_attributes(attributes)
        )
    if name in {"bdr", "shd"}:
        return _border_or_shading_attributes(name, attributes)
    return False


def _integer_in_range(value: str, minimum: int, maximum: int) -> bool:
    value = value.strip(" \t\r\n")
    if not _decimal_integer(value):
        return False
    # Bound significant digits before int(); arbitrary leading zeroes are legal.
    digits = value.lstrip("+-").lstrip("0") or "0"
    if len(digits) > len(str(max(abs(minimum), abs(maximum)))):
        return False
    number = int(digits) * (-1 if value.startswith("-") else 1)
    return minimum <= number <= maximum


def _decimal_integer(value: str) -> bool:
    return re.fullmatch(r"[+-]?[0-9]+", value.strip(" \t\r\n")) is not None


def _unsigned_measure(value: str) -> bool:
    # ECMA-376 Transitional ST_HpsMeasure / ST_TwipsMeasure use unsignedLong,
    # not the narrower application-specific limits of the Office SDK.
    value = value.strip(" \t\r\n")
    return (
        _integer_in_range(value, 0, 2**64 - 1)
        or re.fullmatch(r"[0-9]+(?:\.[0-9]+)?(?:mm|cm|in|pt|pc|pi)", value) is not None
    )


def _font_attributes(attributes: dict[str, str]) -> bool:
    strings = {"ascii", "hAnsi", "eastAsia", "cs"}
    themes = {"asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"}
    return (
        set(attributes).issubset(strings | themes | {"hint"})
        and all(attributes[key] in _THEME_FONTS for key in themes & attributes.keys())
        and attributes.get("hint", "default") in {"default", "eastAsia", "cs"}
    )


def _east_asian_attributes(attributes: dict[str, str]) -> bool:
    toggles = {"combine", "vert", "vertCompress"}
    return (
        set(attributes).issubset(toggles | {"id", "combineBrackets"})
        and all(attributes[key] in _ON_OFF for key in toggles & attributes.keys())
        and ("id" not in attributes or _decimal_integer(attributes["id"]))
        and attributes.get("combineBrackets", "none") in {"none", "round", "square", "angle", "curly"}
    )


def _color_attributes(attributes: dict[str, str]) -> bool:
    return (
        all(
            re.fullmatch(r"auto|[0-9A-Fa-f]{6}", attributes[key]) is not None
            for key in {"color", "fill"} & attributes.keys()
        )
        and all(attributes[key] in _THEME_COLORS for key in {"themeColor", "themeFill"} & attributes.keys())
        and all(
            re.fullmatch(r"[0-9A-Fa-f]{2}", attributes[key]) is not None
            for key in {"themeTint", "themeShade", "themeFillTint", "themeFillShade"} & attributes.keys()
        )
    )


def _border_or_shading_attributes(name: str, attributes: dict[str, str]) -> bool:
    allowed = {"val", "color", "themeColor", "themeTint", "themeShade"}
    if name == "shd":
        allowed |= {"fill", "themeFill", "themeFillTint", "themeFillShade"}
        return (
            set(attributes).issubset(allowed)
            and attributes.get("val") in _SHADING_PATTERNS
            and _color_attributes(attributes)
        )
    allowed |= {"sz", "space", "shadow", "frame"}
    return (
        set(attributes).issubset(allowed)
        and attributes.get("val") in _BORDERS
        and _color_attributes(attributes)
        and all(attributes[key] in _ON_OFF for key in {"shadow", "frame"} & attributes.keys())
        and all(_integer_in_range(attributes[key], 0, 2**64 - 1) for key in {"sz", "space"} & attributes.keys())
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
            or not _control_property_values(child)
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
