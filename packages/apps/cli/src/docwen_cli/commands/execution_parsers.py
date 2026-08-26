"""Protocol 3 execution command tree and argument descriptions."""

from __future__ import annotations

import argparse
from typing import Any

from docwen_cli.commands.execution_options import (
    CHECK_CHOICES,
    HEADING_MERGE_MODE_CHOICES,
    HEADING_NUMBERING_RENDER_MODE_CHOICES,
    IMAGE_LINK_STYLE_CHOICES,
    IMAGE_MODE_CHOICES,
    OCR_LANGUAGE_CHOICES,
    OCR_PLACEMENT_CHOICES,
    TABLE_MERGE_STRATEGY_CHOICES,
)
from docwen_cli.i18n import cli_t
from docwen_cli.parser import bounded_integer, get_common_parser


def _leaf(subparsers: Any, name: str, help_text: str) -> argparse.ArgumentParser:
    return subparsers.add_parser(name, parents=[get_common_parser()], help=help_text)


def _add_timeout(parser: argparse.ArgumentParser, default: int) -> None:
    parser.add_argument(
        "--timeout",
        type=bounded_integer(1, 1800, label="timeout"),
        default=default,
        metavar="SECONDS",
        help=f"Total deadline in seconds (1-1800, default: {default}).",
    )


def _add_admission_control(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--use-detected-format",
        action="store_true",
        help=(
            "Explicitly process high-confidence detected content when its file family "
            "or filename extension conflicts with the declared suffix."
        ),
    )


def _add_write_controls(parser: argparse.ArgumentParser, *, timeout: int = 600) -> None:
    parser.add_argument("--overwrite", action="store_true", help="Replace an existing output target.")
    parser.add_argument("--dry-run", action="store_true", help="Validate without creating output.")
    _add_admission_control(parser)
    _add_timeout(parser, timeout)


def _add_convert_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--template", metavar="ID", help=cli_t("cli.help.template"))
    parser.add_argument("--optimization", metavar="ID")
    parser.add_argument("--check", action="append", choices=sorted(CHECK_CHOICES))
    extraction = parser.add_mutually_exclusive_group()
    extraction.add_argument("--extract-img", action="store_true")
    extraction.add_argument("--no-extract-img", action="store_true")
    parser.add_argument("--ocr", action="store_true")
    parser.add_argument("--ocr-language", choices=sorted(OCR_LANGUAGE_CHOICES))
    parser.add_argument("--image-mode", choices=sorted(IMAGE_MODE_CHOICES))
    parser.add_argument("--image-link-style", choices=sorted(IMAGE_LINK_STYLE_CHOICES))
    parser.add_argument("--table-merge-strategy", choices=sorted(TABLE_MERGE_STRATEGY_CHOICES))
    parser.add_argument("--ocr-placement", choices=sorted(OCR_PLACEMENT_CHOICES))
    parser.add_argument("--clean-numbering", choices=["default", "remove", "keep"])
    parser.add_argument("--add-numbering")
    parser.add_argument("--heading-merge-mode", choices=sorted(HEADING_MERGE_MODE_CHOICES))
    parser.add_argument(
        "--heading-numbering-render-mode",
        choices=sorted(HEADING_NUMBERING_RENDER_MODE_CHOICES),
    )
    parser.add_argument("--spreadsheet-password-prompt", action="store_true")
    parser.add_argument("--allow-spreadsheet-protection-loss", action="store_true")


def register_execution_parsers(subparsers: Any) -> None:
    """Register every protocol 3 execution domain."""

    convert = _leaf(subparsers, "convert", "Convert one file to an explicit destination.")
    convert.add_argument("file")
    convert.add_argument("--to", required=True, metavar="FORMAT")
    convert.add_argument(
        "--output",
        required=True,
        metavar="PATH",
        help="Output file, or the parent directory when --to is Markdown.",
    )
    _add_convert_options(convert)
    _add_write_controls(convert)

    validate = _leaf(
        subparsers,
        "validate",
        "Validate DOCX, Markdown, or legacy Word-family content without changing it.",
    )
    validate.add_argument("file")
    validate.add_argument("--check", action="append", choices=sorted(CHECK_CHOICES))
    validate.add_argument("--report", metavar="PATH")
    _add_admission_control(validate)
    _add_timeout(validate, 30)

    number = _leaf(subparsers, "number", "Normalize document numbering.")
    number_sub = number.add_subparsers(dest="number_command", required=True)
    markdown = _leaf(number_sub, "markdown", "Add or remove Markdown heading numbering.")
    markdown.add_argument("file")
    markdown.add_argument(
        "--operation",
        choices=("add", "remove"),
        required=True,
        help="Explicitly add or remove heading numbering.",
    )
    destination = markdown.add_mutually_exclusive_group(required=True)
    destination.add_argument("--output", metavar="PATH")
    destination.add_argument("--in-place", action="store_true")
    markdown.add_argument(
        "--scheme",
        help="Numbering scheme for --operation add (default: hierarchical_standard).",
    )
    _add_write_controls(markdown)

    merge = _leaf(subparsers, "merge", "Merge multiple inputs into one output.")
    merge_sub = merge.add_subparsers(dest="merge_command", required=True)
    pdf = _leaf(merge_sub, "pdf", "Merge PDF files.")
    pdf.add_argument("files", nargs="+")
    pdf.add_argument("--output", required=True, metavar="PATH")
    _add_write_controls(pdf)
    tables = _leaf(merge_sub, "tables", "Merge spreadsheet tables.")
    tables.add_argument("files", nargs="+")
    tables.add_argument("--output", required=True, metavar="PATH")
    tables.add_argument("--mode")
    _add_write_controls(tables)
    images = _leaf(merge_sub, "images", "Merge images into a TIFF file.")
    images.add_argument("files", nargs="+")
    images.add_argument("--output", required=True, metavar="PATH")
    images.add_argument("--keep-alpha", action="store_true")
    _add_write_controls(images)

    split = _leaf(subparsers, "split", "Split a file into multiple outputs.")
    split_sub = split.add_subparsers(dest="split_command", required=True)
    split_pdf = _leaf(split_sub, "pdf", "Split selected PDF pages.")
    split_pdf.add_argument("file")
    split_pdf.add_argument("--pages", required=True)
    split_pdf.add_argument("--output-dir", required=True, metavar="DIR")
    _add_write_controls(split_pdf)

    batch = _leaf(subparsers, "batch", "Run an explicit multi-file operation.")
    batch_sub = batch.add_subparsers(dest="batch_command", required=True)
    batch_convert = _leaf(batch_sub, "convert", "Convert multiple files to one directory.")
    batch_convert.add_argument("files", nargs="+")
    batch_convert.add_argument("--to", required=True, metavar="FORMAT")
    batch_convert.add_argument("--output-dir", required=True, metavar="DIR")
    _add_convert_options(batch_convert)
    _add_batch_controls(batch_convert)
    batch_validate = _leaf(
        batch_sub,
        "validate",
        "Validate multiple DOCX, Markdown, or legacy Word-family files.",
    )
    batch_validate.add_argument("files", nargs="+")
    batch_validate.add_argument("--check", action="append", choices=sorted(CHECK_CHOICES))
    batch_validate.add_argument("--report-dir", metavar="DIR")
    _add_batch_controls(batch_validate)


def _add_batch_controls(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--jobs", type=bounded_integer(1, 32, label="jobs"), default=1)
    parser.add_argument("--continue-on-error", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    _add_admission_control(parser)
    _add_timeout(parser, 1800)


__all__ = ["register_execution_parsers"]
