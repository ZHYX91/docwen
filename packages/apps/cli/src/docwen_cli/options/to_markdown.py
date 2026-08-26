"""Shared option builder / normalizer for app-level to-markdown conversion.

Single source of truth for to-md option keys.  CLI and GUI both call
these functions so that the options dict always uses the canonical keys
for shared, user-facing Markdown export controls.  Route-specific plugin
options may still live only on their owning manifest.
"""

from __future__ import annotations

# ── Shared canonical key set (mirrors public app controls) ────────────────
# These are the only keys emitted by the shared CLI/GUI to-markdown builder.
# They match the common properties in document / spreadsheet / image / layout
# / presentation manifest option schemas; plugin-specific options remain
# owned by the route that declares them.

CANONICAL_KEYS = frozenset(
    {
        "to_md_keep_images",
        "to_md_enable_ocr",
        "image_mode",
        "ocr_placement",
        "ocr_language",
        "image_link_style",
        "table_merge_strategy",
    }
)

# Exact values admitted by the shared option contract.
_IMAGE_MODE_VALUES = frozenset({"file", "base64", "embed", "omit"})
_OCR_PLACEMENT_VALUES = frozenset({"image_md", "main_md"})
_OCR_LANGUAGE_VALUES = frozenset(
    {"auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"}
)
_LINK_STYLE_VALUES = frozenset({"wiki_embed", "wiki_link", "markdown_embed", "markdown_link"})
_TABLE_MERGE_STRATEGY_VALUES = frozenset({"fill", "empty", "marker"})


def _exact_choice(key: str, value: object, allowed: frozenset[str]) -> str:
    if not isinstance(value, str) or value not in allowed:
        raise ValueError(f"Invalid {key}: {value!r}")
    return value


def _exact_bool(key: str, value: object) -> bool:
    if not isinstance(value, bool):
        raise ValueError(f"Invalid {key}: expected boolean")
    return value


def build_to_markdown_options(
    *,
    keep_images: bool | None = None,
    enable_ocr: bool | None = None,
    image_mode: str | None = None,
    ocr_placement: str | None = None,
    ocr_language: str | None = None,
    image_link_style: str | None = None,
    table_merge_strategy: str | None = None,
) -> dict[str, object]:
    """Build a canonical to-markdown options dict from semantic fields.

    Callers (CLI, GUI) pass their high-level values; the builder
    validate and emit only manifest-conformant keys.  ``None``
    values are omitted so that plugin defaults apply.
    """
    options: dict[str, object] = {}

    if keep_images is not None:
        options["to_md_keep_images"] = _exact_bool("to_md_keep_images", keep_images)

    if enable_ocr is not None:
        options["to_md_enable_ocr"] = _exact_bool("to_md_enable_ocr", enable_ocr)

    if image_mode is not None:
        options["image_mode"] = _exact_choice("image_mode", image_mode, _IMAGE_MODE_VALUES)

    if ocr_placement is not None:
        options["ocr_placement"] = _exact_choice("ocr_placement", ocr_placement, _OCR_PLACEMENT_VALUES)

    if ocr_language is not None:
        options["ocr_language"] = _exact_choice("ocr_language", ocr_language, _OCR_LANGUAGE_VALUES)

    if image_link_style is not None:
        options["image_link_style"] = _exact_choice("image_link_style", image_link_style, _LINK_STYLE_VALUES)

    if table_merge_strategy is not None:
        options["table_merge_strategy"] = _exact_choice(
            "table_merge_strategy", table_merge_strategy, _TABLE_MERGE_STRATEGY_VALUES
        )

    return options


def normalize_to_markdown_options(raw: dict[str, object]) -> dict[str, object]:
    """Validate a raw options dict and return its exact canonical values."""
    cleaned: dict[str, object] = {}

    for key, value in raw.items():
        if key not in CANONICAL_KEYS:
            raise ValueError(f"Unsupported to-Markdown option: {key!r}")

        if key in ("to_md_keep_images", "to_md_enable_ocr"):
            cleaned[key] = _exact_bool(key, value)
        elif key == "image_mode":
            cleaned[key] = _exact_choice(key, value, _IMAGE_MODE_VALUES)
        elif key == "ocr_placement":
            cleaned[key] = _exact_choice(key, value, _OCR_PLACEMENT_VALUES)
        elif key == "ocr_language":
            cleaned[key] = _exact_choice(key, value, _OCR_LANGUAGE_VALUES)
        elif key == "image_link_style":
            cleaned[key] = _exact_choice(key, value, _LINK_STYLE_VALUES)
        elif key == "table_merge_strategy":
            cleaned[key] = _exact_choice(key, value, _TABLE_MERGE_STRATEGY_VALUES)

    return cleaned
