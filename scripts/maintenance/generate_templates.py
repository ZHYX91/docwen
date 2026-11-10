"""Complete localized DOCX distribution templates from their authored assets.

Read locale metadata, retain each existing template's typography, complete the
managed style registry, and apply distribution-owned direction/layout defaults.
New locales start with the English template, falling back to the blank source.
"""

from __future__ import annotations

import argparse
import logging
import shutil
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from docwen_core.docx_styles import MANAGED_DOCUMENT_STYLES, DocumentStyleDefinition
from docwen_plugin_markdown.to_docx.managed_styles import apply_arabic_style_defaults, complete_managed_styles
from docwen_runtime.config.document_styles import build_document_style_catalog

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SCRIPT_DIR = Path(__file__).resolve().parent
LOCALES_DIR = PROJECT_ROOT / "i18n" / "locales"
TEMPLATES_DIR = PROJECT_ROOT / "templates"
BLANK_TEMPLATE_SRC = SCRIPT_DIR / "空白模板.docx"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "samples" / "generated_templates"

logging.basicConfig(level=logging.INFO, format="%(levelname)s - %(name)s - %(message)s")
logger = logging.getLogger(__name__)


def _read_toml(path: Path) -> dict[str, Any]:
    """Read a locale TOML file with the repo's preferred TOML reader."""
    try:
        from docwen_runtime.toml_io import read_toml_file
    except Exception:
        import tomllib

        return tomllib.loads(path.read_text(encoding="utf-8"))
    return read_toml_file(path)


def discover_locale_template_names() -> dict[str, str]:
    """Return ``{locale_code: template_name}`` from locale ``[meta]`` tables."""
    if not LOCALES_DIR.is_dir():
        raise FileNotFoundError(f"Locale directory not found: {LOCALES_DIR}")

    result: dict[str, str] = {}
    for path in sorted(LOCALES_DIR.glob("*.toml"), key=lambda item: item.name.casefold()):
        data = _read_toml(path)
        meta = data.get("meta", {})
        template_name = meta.get("template_name") if isinstance(meta, dict) else None
        if isinstance(template_name, str) and template_name.strip():
            result[path.stem] = template_name.strip()
        else:
            logger.debug("Skipping %s: [meta].template_name is missing", path.name)

    # Multiple interface locales can deliberately point to one shipped default
    # template. Never regenerate a shared filename on behalf of an alias:
    # doing so could overwrite the canonical template's contents.
    template_owners: dict[str, str] = {}
    for locale in sorted(result, key=lambda code: (code != "en_US", code)):
        template_name = result[locale]
        if template_name in template_owners:
            logger.info(
                "Skipping locale %s for template generation: %s reuses %s's template",
                locale,
                template_name,
                template_owners[template_name],
            )
            continue
        template_owners[template_name] = locale

    unique: dict[str, str] = {}
    for locale, template_name in result.items():
        if template_owners[template_name] == locale:
            unique[locale] = template_name
    logger.info("Discovered %d dedicated locale templates: %s", len(unique), ", ".join(unique))
    return unique


def get_locale_placeholders(locale: str) -> dict[str, str]:
    """Return the localized title/body placeholders for *locale*."""
    locale_path = LOCALES_DIR / f"{locale}.toml"
    if not locale_path.is_file():
        raise FileNotFoundError(f"Locale file not found: {locale_path}")

    data = _read_toml(locale_path)
    placeholders = data.get("placeholders", {})
    if not isinstance(placeholders, dict):
        placeholders = {}
    title = str(placeholders.get("title") or "title").strip()
    body = str(placeholders.get("body") or "body").strip()
    return {"title": title or "title", "body": body or "body"}


def _replace_paragraph_text(paragraph: Any, text: str) -> None:
    """Replace paragraph text while keeping the first run's formatting."""
    if paragraph.runs:
        first_run = paragraph.runs[0]
        first_run.text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def _replace_placeholder_text(doc: Any, old_text: str, new_text: str) -> int:
    """Replace placeholder paragraphs/cells and return the replacement count."""
    replacements = 0
    needle = f"{{{{{old_text}}}}}"
    replacement = f"{{{{{new_text}}}}}"

    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == needle:
            _replace_paragraph_text(paragraph, replacement)
            replacements += 1

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text.strip() == needle:
                        _replace_paragraph_text(paragraph, replacement)
                        replacements += 1

    return replacements


def generate_template(locale: str, output_path: Path) -> None:
    """Generate one localized DOCX template."""
    if not BLANK_TEMPLATE_SRC.is_file():
        raise FileNotFoundError(f"Blank template not found: {BLANK_TEMPLATE_SRC}")

    placeholders = get_locale_placeholders(locale)
    template_name = discover_locale_template_names()[locale]
    existing = TEMPLATES_DIR / f"{template_name}.docx"
    # Preserve authored layout in shipped templates; only complete missing styles.
    source = existing if existing.is_file() else TEMPLATES_DIR / "English General Template.docx"
    doc = Document(str(source if source.is_file() else BLANK_TEMPLATE_SRC))

    source_title, source_body = "标题", "正文"
    for locale_file in sorted(LOCALES_DIR.glob("*.toml")):
        data = _read_toml(locale_file)
        if data.get("meta", {}).get("template_name") == source.stem:
            tokens = get_locale_placeholders(locale_file.stem)
            source_title, source_body = tokens["title"], tokens["body"]
            break

    title_count = _replace_placeholder_text(doc, source_title, placeholders["title"])
    body_count = _replace_placeholder_text(doc, source_body, placeholders["body"])

    if title_count == 0:
        raise RuntimeError("Source template does not contain its title placeholder")
    if body_count == 0:
        raise RuntimeError("Source template does not contain its body placeholder")

    catalog = build_document_style_catalog({}, request_options={"locale": locale}, locales_dir=LOCALES_DIR)
    doc, bindings = complete_managed_styles(doc, catalog)
    if locale == "ar_SA":
        for definition in MANAGED_DOCUMENT_STYLES:
            apply_arabic_style_defaults(bindings.get(definition.semantic_key)._element, definition)
        for key, style_id in bindings.foundation_style_ids:
            if key == "normal":
                apply_arabic_style_defaults(
                    next(style for style in doc.styles if style.style_id == style_id)._element,
                    DocumentStyleDefinition("normal", style_id, "paragraph", ""),
                )
        for paragraph in doc.paragraphs:
            is_title = paragraph.text.strip() == f"{{{{{placeholders['title']}}}}}"
            if is_title:
                apply_arabic_style_defaults(
                    paragraph.style._element,
                    DocumentStyleDefinition("title", paragraph.style.style_id, "paragraph", "Normal"),
                )
            p_pr = paragraph._p.get_or_add_pPr()
            bidi = p_pr.find(qn("w:bidi"))
            if bidi is None:
                bidi = OxmlElement("w:bidi")
                p_pr.insert_element_before(bidi, "w:spacing", "w:ind", "w:jc", "w:rPr")
            bidi.set(qn("w:val"), "1")
            # Word's left alignment is the leading edge of an RTL paragraph.
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_title else WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = 0
            for run in paragraph.runs:
                run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cs"), "Arial")

    for definition in MANAGED_DOCUMENT_STYLES:
        if definition.semantic_key.startswith("heading_"):
            heading = bindings.get(definition.semantic_key).paragraph_format
            heading.keep_with_next = True
            heading.space_before = Pt(12)
            heading.space_after = Pt(6)

    if locale not in {"zh_CN", "zh_TW", "ja_JP"}:
        # The inherited CJK grid can double Korean/Russian line height and
        # quantize Western heading spacing. Keep it only for CJK grid layouts.
        for section in doc.sections:
            grid = section._sectPr.find(qn("w:docGrid"))
            if grid is not None:
                grid.set(qn("w:type"), "default")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


def install_templates(output_dir: Path) -> int:
    """Copy generated DOCX templates into the runtime templates directory."""
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    installed = 0
    for src in sorted(output_dir.glob("*.docx"), key=lambda item: item.name.casefold()):
        shutil.copy2(src, TEMPLATES_DIR / src.name)
        logger.info("Installed template: %s", src.name)
        installed += 1
    return installed


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    locale_template_names = discover_locale_template_names()
    parser = argparse.ArgumentParser(
        description="Generate localized DOCX templates from the blank template",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--locale",
        nargs="+",
        choices=sorted(locale_template_names),
        default=sorted(locale_template_names),
        help="Locales to generate (default: all discovered locales)",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Output directory (default: {DEFAULT_OUTPUT_DIR})",
    )
    parser.add_argument(
        "--install",
        action="store_true",
        help="Copy generated templates into templates/ after generation",
    )
    parser.set_defaults(locale_template_names=locale_template_names)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    locale_template_names: dict[str, str] = args.locale_template_names
    output_dir: Path = args.output_dir
    results: dict[str, bool] = {}

    logger.info("=" * 60)
    logger.info("Localized DOCX template generation")
    logger.info("Blank template: %s", BLANK_TEMPLATE_SRC)
    logger.info("Output directory: %s", output_dir)
    logger.info("Locales: %s", ", ".join(args.locale))
    logger.info("=" * 60)

    for locale in args.locale:
        template_name = locale_template_names[locale]
        output_path = output_dir / f"{template_name}.docx"
        try:
            generate_template(locale, output_path)
        except Exception as exc:
            results[locale] = False
            logger.error("%s failed: %s", locale, exc, exc_info=True)
            continue

        results[locale] = True
        logger.info("%s generated: %s (%d bytes)", locale, output_path, output_path.stat().st_size)

    success_count = sum(1 for ok in results.values() if ok)
    fail_count = len(results) - success_count

    if args.install and success_count:
        installed = install_templates(output_dir)
        logger.info("Installed %d templates into %s", installed, TEMPLATES_DIR)

    logger.info("Generation summary: success=%d / %d, failed=%d", success_count, len(results), fail_count)
    return 1 if fail_count else 0


if __name__ == "__main__":
    raise SystemExit(main())
