"""Validation and option projection for protocol 3 execution commands."""

from __future__ import annotations

import argparse
import contextlib
import getpass
from collections.abc import Collection
from typing import Any

from docwen_cli.options.to_markdown import build_to_markdown_options
from docwen_core.text import OCR_LANGUAGE_AUTO, OCR_LANGUAGE_MODELS
from docwen_runtime.templates import is_canonical_template_id

CHECK_CHOICES = frozenset({"punct", "typo", "symbol", "sensitive", "all", "none"})
HEADING_MERGE_MODE_CHOICES = frozenset({"punct_required", "always", "never"})
IMAGE_MODE_CHOICES = frozenset({"file", "base64", "embed", "omit"})
OCR_PLACEMENT_CHOICES = frozenset({"image_md", "main_md"})
OCR_LANGUAGE_CHOICES = frozenset({OCR_LANGUAGE_AUTO, *OCR_LANGUAGE_MODELS.keys()})
IMAGE_LINK_STYLE_CHOICES = frozenset({"wiki_embed", "wiki_link", "markdown_embed", "markdown_link"})
TABLE_MERGE_STRATEGY_CHOICES = frozenset({"fill", "empty", "marker"})
HEADING_NUMBERING_RENDER_MODE_CHOICES = frozenset({"text", "word_native"})


def validate_execution_options(
    args: argparse.Namespace,
) -> None:
    """Validate cross-option value constraints before route projection.

    Source/target applicability belongs exclusively to the resolved Runtime
    route's ``options`` declaration.
    """

    template = getattr(args, "template", None)
    if template and not is_canonical_template_id(str(template)):
        raise ValueError("--template requires an exact canonical ID from resources list templates")

    if getattr(args, "spreadsheet_password_prompt", False) and len(getattr(args, "files", ()) or ()) != 1:
        raise ValueError("--spreadsheet-password-prompt requires exactly one input")

    ocr_language = getattr(args, "ocr_language", None)
    if ocr_language and not getattr(args, "ocr", False):
        raise ValueError("--ocr-language 需要与 --ocr 同时使用")
    image_mode = getattr(args, "image_mode", None)
    if getattr(args, "no_extract_img", False) and image_mode:
        raise ValueError("--image-mode 不能与 --no-extract-img 同时使用")
    ocr_placement = getattr(args, "ocr_placement", None)
    if ocr_placement and not getattr(args, "ocr", False):
        raise ValueError("--ocr-placement 需要与 --ocr 同时使用")
    if ocr_placement and getattr(args, "no_extract_img", False):
        raise ValueError("--ocr-placement 不能与 --no-extract-img 同时使用")
    if str(image_mode or "").strip().lower() == "base64" and str(ocr_placement or "").strip().lower() == "image_md":
        raise ValueError("--image-mode base64 与 --ocr-placement image_md 不能同时使用；请改用 --ocr-placement main_md")
    checks = getattr(args, "check", None) or []
    if "none" in checks and len(checks) > 1:
        raise ValueError("--check none 不能与其它 --check 同时使用")


def parse_pages(value: str) -> list[int]:
    """Parse a comma-separated page/range expression into unique page numbers."""

    pages: list[int] = []
    for part in str(value).split(","):
        part = part.strip()
        if "-" in part:
            start_text, _, end_text = part.partition("-")
            with contextlib.suppress(ValueError):
                start, end = int(start_text), int(end_text)
                if start > end:
                    start, end = end, start
                pages.extend(range(start, end + 1))
        else:
            with contextlib.suppress(ValueError):
                pages.append(int(part))
    return sorted(set(pages))


def normalize_proofread_options(checks: list[str]) -> dict[str, Any]:
    if not checks:
        return {}
    if "none" in checks:
        return {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": False,
            "enable_typos_rule": False,
            "enable_sensitive_word": False,
        }
    enabled = set(checks)
    if "all" in enabled:
        enabled.update({"punct", "typo", "symbol", "sensitive"})
    return {
        "enable_symbol_pairing": "punct" in enabled,
        "enable_symbol_correction": "symbol" in enabled,
        "enable_typos_rule": "typo" in enabled,
        "enable_sensitive_word": "sensitive" in enabled,
    }


def normalize_numbering_options(
    clean_mode: str | None,
    add_mode: str | None,
    render_mode: str | None = None,
) -> dict[str, Any]:
    if clean_mode is None and add_mode is None and render_mode is None:
        return {}
    normalized_clean = str(clean_mode).strip().lower() if clean_mode is not None else "default"
    normalized_add = str(add_mode).strip().lower() if add_mode is not None else "default"
    if normalized_clean not in {"default", "remove", "keep"}:
        raise ValueError(f"清理序号模式不合法: {normalized_clean}")

    add_numbering = normalized_add not in {"default", "none"}
    result: dict[str, Any] = {
        "remove_numbering": normalized_clean == "remove",
        "add_numbering": add_numbering,
        "numbering_scheme": normalized_add if add_numbering else "",
    }
    if render_mode is not None:
        normalized_render = str(render_mode).strip().lower()
        if normalized_render not in {"text", "word_native"}:
            raise ValueError(f"标题序号渲染模式不合法: {render_mode}")
        result["heading_numbering_render_mode"] = normalized_render
    return result


def build_execution_options(
    args: argparse.Namespace,
    *,
    route_options: Collection[str] | None = None,
) -> dict[str, Any]:
    """Normalize explicit CLI values without guessing route applicability."""

    options: dict[str, Any] = {}

    if template := getattr(args, "template", None):
        options["template_name"] = str(template)
    options.update(normalize_proofread_options(getattr(args, "check", None) or []))

    keep_images = True if getattr(args, "extract_img", False) else None
    if getattr(args, "no_extract_img", False):
        keep_images = False
    options.update(
        build_to_markdown_options(
            keep_images=keep_images,
            enable_ocr=True if getattr(args, "ocr", False) else None,
            image_mode=getattr(args, "image_mode", None),
            ocr_placement=getattr(args, "ocr_placement", None),
            ocr_language=getattr(args, "ocr_language", None),
            image_link_style=getattr(args, "image_link_style", None),
            table_merge_strategy=getattr(args, "table_merge_strategy", None),
        )
    )

    options.update(
        normalize_numbering_options(
            getattr(args, "clean_numbering", None),
            getattr(args, "add_numbering", None),
            getattr(args, "heading_numbering_render_mode", None),
        )
    )
    if heading_mode := getattr(args, "heading_merge_mode", None):
        options["heading_merge_mode"] = str(heading_mode).strip().lower()

    if getattr(args, "spreadsheet_password_prompt", False):
        if route_options is not None and "spreadsheet_password" not in route_options:
            raise ValueError("Canonical runtime route does not declare option: spreadsheet_password")
        password = getpass.getpass("Spreadsheet protection password: ")
        if password:
            options["spreadsheet_password"] = password
    if getattr(args, "allow_spreadsheet_protection_loss", False):
        options["allow_spreadsheet_protection_loss"] = True

    if pages := getattr(args, "pages", None):
        options["pages"] = parse_pages(str(pages))
    if dpi := getattr(args, "dpi", None):
        options["render_dpi"] = int(dpi)
    if mode := getattr(args, "mode", None):
        options["merge_mode"] = str(mode)
    if getattr(args, "keep_alpha", False):
        options["keep_alpha"] = True
    return options


__all__ = [
    "CHECK_CHOICES",
    "HEADING_MERGE_MODE_CHOICES",
    "HEADING_NUMBERING_RENDER_MODE_CHOICES",
    "IMAGE_LINK_STYLE_CHOICES",
    "IMAGE_MODE_CHOICES",
    "OCR_LANGUAGE_CHOICES",
    "OCR_PLACEMENT_CHOICES",
    "TABLE_MERGE_STRATEGY_CHOICES",
    "build_execution_options",
    "normalize_numbering_options",
    "normalize_proofread_options",
    "parse_pages",
    "validate_execution_options",
]
