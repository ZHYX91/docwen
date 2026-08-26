from __future__ import annotations

import argparse
import base64
import contextlib
import hashlib
import io
import json
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from bisect import bisect_right
from collections.abc import Sequence
from hashlib import sha256
from pathlib import Path
from typing import IO, Any
from xml.etree import ElementTree

import openpyxl

try:
    from scripts.release import packaged_resources as _packaged_resources
except ModuleNotFoundError:
    import packaged_resources as _packaged_resources  # type: ignore[no-redef]

_REQUIRED_CONFIG_FILES = _packaged_resources.REQUIRED_CONFIG_FILES
_REQUIRED_LOCALE_FILES = _packaged_resources.REQUIRED_LOCALE_FILES
_REQUIRED_MODEL_FILES = _packaged_resources.REQUIRED_MODEL_FILES
_REQUIRED_TEMPLATE_FILES = _packaged_resources.REQUIRED_TEMPLATE_FILES
verify_common_resource_layout = _packaged_resources.verify_common_resource_layout

_PLUGIN_LOAD_FAILURE_MARKER = "Failed to load plugin "
_PYMUPDF_LAYOUT_GATE_ID = "python.pymupdf4llm"
_PYMUPDF_LAYOUT_SMOKE_TEXT = "DOCWEN PACKAGED PYMUPDF LAYOUT SMOKE"
_TEMPLATE_ID_PATTERN = re.compile(r"^template\.(?:docx|xlsx)\.[0-9a-f]{64}$")
_TEMPLATE_RESOURCE_FIELDS = frozenset({"id", "name", "target", "description", "path", "size_bytes", "modified_ns"})
_TEMPLATE_SMOKE_TEXT = "DOCWEN PACKAGED CANONICAL TEMPLATE ID SMOKE"
_PROOFREAD_REPORT_FIXTURE_TEXT = "\ufeff\r\n# 坐标 😀e\u0301👩\u200d💻１２\r\n结尾（"
_LONG_PATH_MINIMUM_LENGTH = 201
_LONG_PATH_SEGMENT_LIMIT = 32
_LONG_PATH_OUTPUT_NAME = "转换 结果.md"
_PROOFREAD_LOCATION_CONTRACT = {
    "id": "docwen.proofread-text-range",
    "version": 1,
    "coordinate_system": "unicode_code_point",
    "offset_base": 0,
    "line_base": 0,
    "column_base": 0,
    "range_end": "exclusive",
}
_DOCTOR_BASE_CHECK_IDS = frozenset(
    {"config.load", "path.temp_directory", "security.dependency_egress_guard"},
)
_WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_WORD_TAG = f"{{{_WORDPROCESSINGML_NAMESPACE}}}"
MACHINE_DOCUMENT_SEMANTICS_LIMITATIONS = (
    {
        "severity": "warning",
        "code": "document_semantics.citation_processor_unavailable",
        "message": (
            "DocWen does not run a CSL citation processor or accept citation_style inputs in Machine v1; "
            "Markdown citation keys remain literal."
        ),
    },
    {
        "severity": "warning",
        "code": "document_semantics.v1_scope",
        "message": (
            "Document semantics v1 excludes CSL processing, composite or range citation semantics, "
            "custom citation display, and PDF semantic round trips."
        ),
    },
)
MACHINE_RESOLVED_DOCUMENT_LIMITATIONS: tuple[dict[str, str], ...] = (
    {
        "severity": "warning",
        "code": "resolved_document.provider_owned_semantics",
        "message": (
            "DocWen consumes already-resolved targets, citations, resources, and numbering facts; "
            "it does not scan a Workspace, run a citation resolver, or infer numbering from authored text."
        ),
    },
)
MACHINE_SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE = "application/vnd.docwen.semantic-bibliography+json"
MACHINE_RESOLVED_DOCUMENT_MEDIA_TYPE = "application/vnd.docwen.resolved-document+json"
MACHINE_NUMBERING_EXPORT_PLAN_MEDIA_TYPE = "application/vnd.docwen.numbering-export-plan+json"
MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT = {
    "$schema": "urn:docwen:schema:resolved-document:v1",
    "schema": "docwen.resolved_document.v1",
    "input_id": "fixture-1d28a20836a07c7818a289b51a7ef4ca",
    "source_sha256": "14923b0fd8c0e37641aaa761786ead494d5fda46d18d53704634e1010a7a7db0",
    "plan_sha256": "2aecfd6cfa60bf788eecae7fbb7494642fff79af5988969ad34926f0e2425054",
    "document": {
        "authored_markdown": (
            "# Architecture ^h-7f3a\n\n"
            "Figure: System overview ^system-overview\n\n"
            "![[system.png]]\n\n"
            "Table: Results ^results-main\n\n"
            "| Metric | Value |\n|---|---|\n| Score | 95 |\n\n"
            "Equation: ^energy-main\n\n"
            "$$\nE = mc^2\n$$\n\n"
            "Code: Entry point ^entry-main\n\n"
            "```rust\nfn main() {}\n```\n\n"
            "Stable: @[[#^h-7f3a]] and @[[#^system-overview|System overview]].\n"
            "Ordinary: [[#^system-overview]] and ![[Guide#^h-7f3a]].\n"
            "Citation: @cite-one.\n"
            "\n"
            "Notes: default[^alpha], explicit[^footnote:beta], first endnote[^endnote:omega], "
            "second endnote[^endnote:old].\n\n"
            "[^alpha]: Default footnote.\n"
            "[^footnote:beta]: Explicit footnote.\n"
            "[^endnote:omega]: Canonical endnote.\n"
            "[^endnote:old]: Second endnote.\n"
        ),
        "targets": [
            {
                "source_start": 0,
                "source_end": 22,
                "source_slice_sha256": "9db146b8c77e8ad9379cfb192290d6668281b24e680e3228786989beca20265a",
                "kind": "heading",
                "target_id": "h-7f3a",
                "heading_level": 1,
                "authored_text": "Architecture",
            },
            {
                "source_start": 24,
                "source_end": 64,
                "source_slice_sha256": "1277fbdf915dde6cda3ab80a506db9d8ec8a077d9bec04646748fb7369d1f033",
                "kind": "figure",
                "target_id": "system-overview",
                "heading_level": None,
                "authored_text": "System overview",
            },
            {
                "source_start": 83,
                "source_end": 111,
                "source_slice_sha256": "82c070f6467923aa51bc75b9fa5aed5eb7f99232b19b099657a9aa2f9737fcc1",
                "kind": "table",
                "target_id": "results-main",
                "heading_level": None,
                "authored_text": "Results",
            },
            {
                "source_start": 158,
                "source_end": 180,
                "source_slice_sha256": "52484aa20fbf4788aebc2f7343dd7b4b375711b4b720210b0e6ca802cc07297a",
                "kind": "equation",
                "target_id": "energy-main",
                "heading_level": None,
                "authored_text": "",
            },
            {
                "source_start": 198,
                "source_end": 227,
                "source_slice_sha256": "2a756b8a621ec190b9eea85a367a410eb5cd0654a301c31d50f999490a87694a",
                "kind": "code_block",
                "target_id": "entry-main",
                "heading_level": None,
                "authored_text": "Entry point",
            },
        ],
        "references": [
            {
                "source_start": 263,
                "source_end": 276,
                "source_slice_sha256": "850ccceb2327d91136ce996a3577d080a287988504857907e32f679a3819959c",
                "authored_token": "@[[#^h-7f3a]]",
                "target_source_start": 0,
                "target_source_end": 22,
                "target_kind": "heading",
                "target_id": "h-7f3a",
                "cached_number": "1",
                "alias": None,
            },
            {
                "source_start": 281,
                "source_end": 319,
                "source_slice_sha256": "391f374d10a10a9deb0bf5306378e6a55e271355a3a9b8128394c1bd947976ee",
                "authored_token": "@[[#^system-overview|System overview]]",
                "target_source_start": 24,
                "target_source_end": 64,
                "target_kind": "figure",
                "target_id": "system-overview",
                "cached_number": "1",
                "alias": "System overview",
            },
        ],
        "resource_occurrences": [
            {
                "source_start": 66,
                "source_end": 81,
                "source_slice_sha256": "b3709c74910dd4adde5211af6f88b605c25f36a2e74817fead91a52462e9bf4d",
                "authored_token": "![[system.png]]",
                "authored_locator": "system.png",
                "resource_id": "image-system",
            }
        ],
        "citations": [
            {
                "source_start": 387,
                "source_end": 396,
                "source_slice_sha256": "5f6357be6410398a73822925e446d175effed3cdab0d5d5ad38e28ad262eae28",
                "authored_token": "@cite-one",
                "form": "narrative",
                "cluster_id": "citation-one",
                "items": [
                    {
                        "citation_key": "cite-one",
                        "record_id": "ref-one",
                        "record_sha256": "b457bbe0bf6e46d3a84c2e7b02afee45d211e6567e53658c2b37d343013a6df7",
                        "presentation": "One (2026)",
                    }
                ],
                "cached_result": "One (2026)",
            }
        ],
        "resources": [
            {
                "resource_id": "bibliography-main",
                "role": "bibliography",
                "media_type": MACHINE_SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE,
                "size_bytes": 177,
                "sha256": "27bfb40d2fb8044bcf7b5608645da0a95296e1f03920b5317c3967e7adfe0a6c",
                "content_base64": (
                    "eyJzY2hlbWEiOiJkb2N3ZW4uc2VtYW50aWNfYmlibGlvZ3JhcGh5LnYxIiwiZW50cmllcyI6W3siaXRlbV9p"
                    "ZCI6InJlZi1vbmUiLCJydW5zIjpbeyJ0ZXh0IjoiT25lLCBBLiAoMjAyNikuIEV4YWN0IHR3by4iLCJpdGFs"
                    "aWMiOnRydWUsImhyZWYiOiJodHRwczovL2V4YW1wbGUudGVzdC9yZWYtb25lIn1dfV19"
                ),
            },
            {
                "resource_id": "image-system",
                "role": "linked_resource",
                "media_type": "image/png",
                "size_bytes": 75,
                "sha256": "76f04b5505024667600de38630b5595fcbbf7449ccd5ab89baaeb9b4b08f4434",
                "content_base64": (
                    "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNUSFjAwMDAxAAGAA0qASTlOPBg"
                    "AAAAAElFTkSuQmCC"
                ),
            },
        ],
    },
}
MACHINE_EXACT_TWO_NUMBERING_PLAN = {
    "$schema": "urn:docwen:schema:numbering-export-plan:v1",
    "schema": "docwen.numbering_export_plan.v1",
    "input_id": "fixture-1d28a20836a07c7818a289b51a7ef4ca",
    "source_sha256": "14923b0fd8c0e37641aaa761786ead494d5fda46d18d53704634e1010a7a7db0",
    "plan_sha256": "2aecfd6cfa60bf788eecae7fbb7494642fff79af5988969ad34926f0e2425054",
    "plan": {
        "heading_definitions": [
            {
                "definition_id": "heading-default",
                "levels": [
                    {
                        "display": [{"counter": {"level": 1, "number_format": "arabic_half"}}],
                        "level": 1,
                        "number_format": "arabic_half",
                        "restart_after_level": None,
                        "start": 1,
                        "suffix": "space",
                    }
                ],
            }
        ],
        "heading_instances": [{"definition_id": "heading-default", "instance_id": "heading-instance-1", "starts": []}],
        "targets": [
            {
                "derived_number": "1",
                "enabled": True,
                "kind": "heading",
                "materialization": {
                    "definition_id": "heading-default",
                    "instance_id": "heading-instance-1",
                    "level": 1,
                    "type": "heading_list",
                },
                "source_end": 22,
                "source_start": 0,
                "target_id": "h-7f3a",
            },
            {
                "derived_number": "1",
                "enabled": True,
                "kind": "figure",
                "materialization": {
                    "chapter_cached_number": None,
                    "chapter_heading_level": None,
                    "chapter_heading_style": None,
                    "chapter_separator": None,
                    "counter": "Figure",
                    "label_separator": " ",
                    "localized_label": "Figure",
                    "number_format": "arabic_half",
                    "restart_heading_level": None,
                    "restart_heading_style": None,
                    "sequence_action": "continue",
                    "sequence_cached_number": "1",
                    "start_value": None,
                    "type": "simple_seq",
                },
                "source_end": 64,
                "source_start": 24,
                "target_id": "system-overview",
            },
            {
                "derived_number": "1",
                "enabled": True,
                "kind": "table",
                "materialization": {
                    "chapter_cached_number": None,
                    "chapter_heading_level": None,
                    "chapter_heading_style": None,
                    "chapter_separator": None,
                    "counter": "Table",
                    "label_separator": " ",
                    "localized_label": "Table",
                    "number_format": "arabic_half",
                    "restart_heading_level": None,
                    "restart_heading_style": None,
                    "sequence_action": "continue",
                    "sequence_cached_number": "1",
                    "start_value": None,
                    "type": "simple_seq",
                },
                "source_end": 111,
                "source_start": 83,
                "target_id": "results-main",
            },
            {
                "derived_number": "1",
                "enabled": True,
                "kind": "equation",
                "materialization": {
                    "chapter_cached_number": None,
                    "chapter_heading_level": None,
                    "chapter_heading_style": None,
                    "chapter_separator": None,
                    "counter": "Equation",
                    "label_separator": " ",
                    "localized_label": "Equation",
                    "number_format": "arabic_half",
                    "restart_heading_level": None,
                    "restart_heading_style": None,
                    "sequence_action": "continue",
                    "sequence_cached_number": "1",
                    "start_value": None,
                    "type": "simple_seq",
                },
                "source_end": 180,
                "source_start": 158,
                "target_id": "energy-main",
            },
            {
                "derived_number": "1",
                "enabled": True,
                "kind": "code_block",
                "materialization": {
                    "chapter_cached_number": None,
                    "chapter_heading_level": None,
                    "chapter_heading_style": None,
                    "chapter_separator": None,
                    "counter": "Code",
                    "label_separator": " ",
                    "localized_label": "Code",
                    "number_format": "arabic_half",
                    "restart_heading_level": None,
                    "restart_heading_style": None,
                    "sequence_action": "continue",
                    "sequence_cached_number": "1",
                    "start_value": None,
                    "type": "simple_seq",
                },
                "source_end": 227,
                "source_start": 198,
                "target_id": "entry-main",
            },
        ],
    },
}

MACHINE_SEMANTIC_BIBLIOGRAPHY_URL = "https://example.org/neutral-documents"
MACHINE_EXACT_TWO_IMAGE_BYTES = base64.b64decode(
    next(
        resource["content_base64"]
        for resource in MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT["document"]["resources"]
        if resource["role"] == "linked_resource"
    )
)
MACHINE_SEMANTIC_BIBLIOGRAPHY = {
    "schema": "docwen.semantic_bibliography.v1",
    "entries": [
        {
            "item_id": "smith2025",
            "runs": [
                {"text": "Smith, A. ", "bold": True},
                {"text": "Neutral documents", "italic": True, "href": MACHINE_SEMANTIC_BIBLIOGRAPHY_URL},
            ],
        }
    ],
}
MACHINE_PHYSICAL_PAGE_LIMITATIONS = (
    {
        "severity": "warning",
        "code": "physical_page_ocr.best_effort",
        "message": (
            "When OCR is enabled it is best effort: every physical page or frame retains an ordered fragment and "
            "typed status even when recognition is blank or unavailable."
        ),
    },
    {
        "severity": "warning",
        "code": "physical_page_ocr.consumer_owned_import",
        "message": (
            "The Bundle reports page and resource facts only; Node layout, basenames, and import strategy remain "
            "consumer-owned."
        ),
    },
)
_OFD_FIXTURE_SCRIPT = r"""
import contextlib
import io
import sys
import warnings
from pathlib import Path

from loguru import logger as loguru_logger
from PIL import Image, ImageDraw, ImageFont

warnings.filterwarnings("ignore", category=SyntaxWarning)
loguru_logger.disable("easyofd")
from easyofd import OFD

pages = []
try:
    for page_number in range(1, 3):
        page = Image.new("RGB", (900, 1200), "white")
        draw = ImageDraw.Draw(page)
        try:
            font = ImageFont.truetype("DejaVuSans.ttf", 48)
        except OSError:
            font = ImageFont.load_default()
        draw.text((72, 180), f"DOCWEN OFD PAGE {page_number}", fill="black", font=font)
        pages.append(page)
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        payload = OFD().jpg2ofd(pages)
    Path(sys.argv[1]).write_bytes(payload)
finally:
    for page in pages:
        page.close()
"""
MACHINE_DOCUMENT_SEMANTICS_FIXTURE = """# Packaged Machine Protocol

Table: Sales channels {#tbl-sales}

| Region | Sales | < | Total |
|---|---:|---:|---:|
| ^ | Online | Retail | ^ |
| North | 10 | 12 | 22 |
| South | 8 | 9 | 17 |
{header-rows=2 header-cols=1 repeat-header=true}

See @tbl-sales for the synthetic totals.

Existing studies support this result [@smith2025; @wang2024].

![Machine semantic image](assets/机器协议-语义.png)

Machine semantic tail.
"""


def _default_binary_name() -> str:
    return "DocWenCLI.exe" if os.name == "nt" else "DocWenCLI"


def _verify_resource_layout(binary_dir: Path) -> None:
    verify_common_resource_layout(binary_dir, error_prefix="packaged_cli")


def _build_long_path_work_dir(temp_root: Path) -> Path:
    """Build the same 200+ character Unicode fixture from short or long temp roots."""

    work_dir = temp_root / "资料 空格"
    for index in range(_LONG_PATH_SEGMENT_LIMIT):
        if len(str(work_dir / _LONG_PATH_OUTPUT_NAME)) >= _LONG_PATH_MINIMUM_LENGTH:
            return work_dir
        work_dir /= f"{index:02d}-长路径验证目录长路径验证目录长路径验证目录"
    raise RuntimeError(f"packaged_cli_long_path_fixture_too_short: {work_dir / _LONG_PATH_OUTPUT_NAME}")


def _write_xlsx(path: Path) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    if worksheet is None:
        raise RuntimeError("packaged_cli_workbook_has_no_active_sheet")
    worksheet["A1"] = "name"
    worksheet["B1"] = "value"
    worksheet["A2"] = "alpha"
    worksheet["B2"] = 1
    workbook.save(path)
    workbook.close()


def _write_pymupdf_layout_pdf(path: Path) -> None:
    """Create a tiny, deterministic PDF that exercises the packaged layout route."""
    import fitz

    document = fitz.open()
    try:
        page = document.new_page(width=480, height=180)
        page.insert_text((40, 90), _PYMUPDF_LAYOUT_SMOKE_TEXT, fontsize=16)
        document.save(path)
    finally:
        document.close()


def _write_physical_page_pdf(path: Path, *, copy_window_bytes: int = 16 * 1024 * 1024) -> None:
    """Create the canonical P=4/K=4 portion of the packaged physical-page corpus."""

    import fitz
    from PIL import Image, ImageDraw, ImageFont

    document = fitz.open()
    try:
        for page_number in range(1, 5):
            image = Image.new("RGB", (900, 240), "white")
            draw = ImageDraw.Draw(image)
            try:
                font = ImageFont.truetype("DejaVuSans.ttf", 48)
            except OSError:
                font = ImageFont.load_default()
            if page_number != 2:
                draw.text((36, 72), f"DOCWEN PHYSICAL PAGE {page_number}", fill="black", font=font)
            with tempfile.SpooledTemporaryFile(max_size=2 * 1024 * 1024) as encoded:
                image.save(encoded, format="PNG")
                encoded.seek(0)
                page = (
                    document.new_page(width=14400, height=4)
                    if page_number == 3
                    else document.new_page(width=480, height=180)
                )
                page.insert_image(page.rect, stream=encoded.read())
            image.close()
        if copy_window_bytes:
            document.embfile_add("copy-window.bin", random.Random(2404).randbytes(copy_window_bytes))
        document.save(path)
    finally:
        document.close()


def _write_physical_page_tiff(path: Path) -> None:
    """Create a deterministic four-frame TIFF for the packaged Machine gate."""

    from PIL import Image, ImageDraw, ImageFont

    frames = []
    try:
        for page_number in range(1, 5):
            frame = Image.new("RGB", (65535, 1), "white") if page_number == 3 else Image.new("RGB", (900, 240), "white")
            draw = ImageDraw.Draw(frame)
            try:
                font = ImageFont.truetype("DejaVuSans.ttf", 48)
            except OSError:
                font = ImageFont.load_default()
            if page_number not in {2, 3}:
                draw.text((36, 72), f"DOCWEN TIFF FRAME {page_number}", fill="black", font=font)
            frames.append(frame)
        frames[0].save(path, format="TIFF", save_all=True, append_images=frames[1:])
    finally:
        for frame in frames:
            frame.close()


def _write_physical_page_ofd(path: Path) -> None:
    """Create a real two-page OFD using the packaged producer dependency."""

    try:
        # easyofd writes a nested ``./test/Doc_0/...`` tree and redirects
        # process-global stdio. Isolate both behaviors in a child process whose
        # cwd is a short system-temp path; the verifier's parent cwd is never
        # changed, so parallel verifier helpers cannot collide.
        with tempfile.TemporaryDirectory(prefix="dw-ofd-") as generation_dir:
            fixture_name = "fixture.ofd"
            completed = subprocess.run(
                [sys.executable, "-c", _OFD_FIXTURE_SCRIPT, fixture_name],
                cwd=generation_dir,
                capture_output=True,
                timeout=60,
                check=False,
            )
            if completed.returncode != 0 or completed.stdout or completed.stderr:
                raise RuntimeError(
                    "packaged_physical_page_ofd_fixture_generation_failed:"
                    f"{completed.returncode}:{completed.stdout[:500]!r}:{completed.stderr[:500]!r}"
                )
            payload = _read_bytes_with_long_path(Path(generation_dir) / fixture_name)
            with zipfile.ZipFile(io.BytesIO(payload)) as archive:
                required_members = {
                    "OFD.xml",
                    "Doc_0/Pages/Page_0/Content.xml",
                    "Doc_0/Pages/Page_1/Content.xml",
                }
                if not required_members.issubset(archive.namelist()):
                    raise RuntimeError("packaged_physical_page_ofd_fixture_invalid")
    except (OSError, subprocess.SubprocessError, zipfile.BadZipFile) as exc:
        raise RuntimeError("packaged_physical_page_ofd_fixture_generation_failed") from exc
    path.write_bytes(payload)


def _write_physical_page_xps(path: Path) -> None:
    """Create a font-free real two-page XPS with one raster per page."""

    from PIL import Image, ImageDraw, ImageFont

    page_images: dict[str, bytes] = {}
    for page_number in range(1, 3):
        image = Image.new("RGB", (640, 320), "white")
        try:
            draw = ImageDraw.Draw(image)
            try:
                font = ImageFont.truetype("DejaVuSans.ttf", 36)
            except OSError:
                font = ImageFont.load_default()
            draw.text((48, 128), f"DOCWEN XPS PAGE {page_number}", fill="black", font=font)
            encoded = io.BytesIO()
            image.save(encoded, format="PNG")
            page_images[f"Resources/Images/page-{page_number}.png"] = encoded.getvalue()
        finally:
            image.close()
    content_types = """<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/FixedDocumentSequence.fdseq" ContentType="application/vnd.ms-package.xps-fixeddocumentsequence+xml"/>
  <Override PartName="/Documents/1/FixedDocument.fdoc" ContentType="application/vnd.ms-package.xps-fixeddocument+xml"/>
  <Override PartName="/Documents/1/Pages/1.fpage" ContentType="application/vnd.ms-package.xps-fixedpage+xml"/>
  <Override PartName="/Documents/1/Pages/2.fpage" ContentType="application/vnd.ms-package.xps-fixedpage+xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png" ContentType="image/png"/>
</Types>"""
    parts = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": """<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="R1" Type="http://schemas.microsoft.com/xps/2005/06/fixedrepresentation" Target="/FixedDocumentSequence.fdseq"/>
</Relationships>""",
        "FixedDocumentSequence.fdseq": """<FixedDocumentSequence xmlns="http://schemas.microsoft.com/xps/2005/06">
  <DocumentReference Source="Documents/1/FixedDocument.fdoc"/>
</FixedDocumentSequence>""",
        "Documents/1/FixedDocument.fdoc": """<FixedDocument xmlns="http://schemas.microsoft.com/xps/2005/06">
  <PageContent Source="Pages/1.fpage"/>
  <PageContent Source="Pages/2.fpage"/>
</FixedDocument>""",
    }
    for page_number in range(1, 3):
        parts[
            f"Documents/1/Pages/{page_number}.fpage"
        ] = f"""<FixedPage xmlns="http://schemas.microsoft.com/xps/2005/06" Width="816" Height="1056" xml:lang="en-US">
  <Path Data="M88,220 L728,220 728,540 88,540 Z">
    <Path.Fill>
      <ImageBrush ImageSource="/Resources/Images/page-{page_number}.png" Viewbox="0,0,640,320"
                  ViewboxUnits="Absolute" Viewport="88,220,640,320" ViewportUnits="Absolute" TileMode="None"/>
    </Path.Fill>
  </Path>
</FixedPage>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, body in parts.items():
            archive.writestr(name, body)
        for name, body in page_images.items():
            archive.writestr(name, body)


def _verify_physical_page_bundle(
    *,
    terminal: dict[str, Any],
    task_id: str,
    staging_root: Path,
    page_count: int,
    resource_count: int,
    ocr_enabled: bool,
    keep_images: bool,
    expected_statuses: tuple[str, ...] | None = None,
) -> dict[str, Any]:
    """Verify page facts without deriving P from the number of extracted resources."""

    if terminal.get("method") != "task/completed":
        raise RuntimeError(f"packaged_physical_page_terminal_invalid:{terminal}")
    params = terminal.get("params")
    bundle = params.get("bundle") if isinstance(params, dict) else None
    diagnostics = params.get("diagnostics") if isinstance(params, dict) else None
    if not isinstance(bundle, dict) or bundle.get("task_id") != task_id or not isinstance(diagnostics, list):
        raise RuntimeError(f"packaged_physical_page_bundle_invalid:{terminal}")
    artifacts = bundle.get("artifacts")
    relations = bundle.get("relations")
    entries = bundle.get("entries")
    if not isinstance(artifacts, list) or not isinstance(relations, list) or not isinstance(entries, list):
        raise RuntimeError(f"packaged_physical_page_bundle_invalid:{bundle}")
    by_id = {
        artifact.get("artifact_id"): artifact
        for artifact in artifacts
        if isinstance(artifact, dict) and isinstance(artifact.get("artifact_id"), str)
    }
    manifest_relations = [
        relation
        for relation in relations
        if isinstance(relation, dict) and relation.get("type") == "resource_of" and relation.get("role") == "manifest"
    ]
    manifest_ids = {
        relation.get("source_artifact_id")
        for relation in manifest_relations
        if isinstance(relation.get("source_artifact_id"), str)
    }
    if bundle.get("layout_schema") == "docwen.document_node.v1":
        manifest_artifacts = [by_id[artifact_id] for artifact_id in manifest_ids if artifact_id in by_id]
        if (
            len(manifest_relations) != 1
            or len(manifest_artifacts) != 1
            or manifest_artifacts[0].get("media_type") != "application/vnd.docwen.document-node+json"
            or manifest_artifacts[0].get("suggested_name") != "docwen-node.json"
        ):
            raise RuntimeError(f"packaged_document_node_manifest_invalid:{bundle}")
    elif manifest_relations:
        raise RuntimeError(f"packaged_artifact_layout_manifest_invalid:{bundle}")
    deliverable_artifacts = [
        artifact
        for artifact in artifacts
        if isinstance(artifact, dict) and artifact.get("artifact_id") not in manifest_ids
    ]
    semantic_relations = [relation for relation in relations if relation not in manifest_relations]
    documents = [artifact for artifact in deliverable_artifacts if artifact.get("kind") == "document"]
    fragments = [artifact for artifact in deliverable_artifacts if artifact.get("kind") == "fragment"]
    resources = [artifact for artifact in deliverable_artifacts if artifact.get("kind") == "resource"]
    expected_fragment_count = page_count if ocr_enabled else 0
    expected_resource_count = resource_count if keep_images else 0
    if (
        len(documents) != 1
        or len(fragments) != expected_fragment_count
        or len(resources) != expected_resource_count
        or len(deliverable_artifacts) != 1 + expected_fragment_count + expected_resource_count
        or entries
        != [{"artifact_id": documents[0].get("artifact_id"), "role": "primary", "ordinal": 0, "preferred": True}]
    ):
        raise RuntimeError(f"packaged_physical_page_shape_invalid:{bundle}")
    primary_id = documents[0]["artifact_id"]
    page_relations = [
        relation
        for relation in semantic_relations
        if isinstance(relation, dict) and relation.get("type") == "fragment_of" and relation.get("role") == "ocr_page"
    ]
    resource_relations = [
        relation
        for relation in semantic_relations
        if isinstance(relation, dict) and relation.get("type") == "resource_of"
    ]
    if (
        len(page_relations) != expected_fragment_count
        or len(resource_relations) != expected_resource_count
        or len(semantic_relations) != expected_fragment_count + expected_resource_count
        or sorted(str(relation.get("source_artifact_id")) for relation in page_relations)
        != sorted(fragment["artifact_id"] for fragment in fragments)
        or sorted(str(relation.get("source_artifact_id")) for relation in resource_relations)
        != sorted(resource["artifact_id"] for resource in resources)
    ):
        raise RuntimeError(f"packaged_physical_page_relations_invalid:{relations}")
    statuses = tuple(
        relation.get("page_fragment", {}).get("ocr_status")
        for relation in page_relations
        if isinstance(relation.get("page_fragment"), dict)
    )
    if ocr_enabled:
        expected_page_keys = {"fragment_kind", "page_index", "page_count", "ocr_status", "source_page"}
        if any(
            set(relation.get("page_fragment", {})) != expected_page_keys
            or relation["page_fragment"].get("fragment_kind") != "page"
            or any(
                type(relation["page_fragment"].get(key)) is not int
                for key in ("page_index", "page_count", "source_page")
            )
            for relation in page_relations
        ):
            raise RuntimeError(f"packaged_physical_page_fragment_payload_invalid:{page_relations}")
        if (
            [relation["page_fragment"]["page_index"] for relation in page_relations] != list(range(1, page_count + 1))
            or [relation["page_fragment"]["source_page"] for relation in page_relations]
            != list(range(1, page_count + 1))
            or [relation["page_fragment"]["page_count"] for relation in page_relations] != [page_count] * page_count
            or [relation.get("ordinal") for relation in page_relations] != list(range(page_count))
            or any(relation.get("target_artifact_id") != primary_id for relation in page_relations)
        ):
            raise RuntimeError(f"packaged_physical_page_sequence_invalid:{page_relations}")
        allowed_statuses = {
            "success",
            "no_text",
            "input_missing",
            "unavailable",
            "model_missing",
            "initialization_failed",
            "recognition_failed",
        }
        if len(statuses) != page_count or any(status not in allowed_statuses for status in statuses):
            raise RuntimeError(f"packaged_physical_page_status_invalid:{statuses}")
        if expected_statuses is not None and statuses != expected_statuses:
            raise RuntimeError(f"packaged_physical_page_status_mismatch:{statuses}")
        primary_payload = _read_bytes_with_long_path(staging_root / Path(documents[0]["locator"]))
        for relation, status in zip(page_relations, statuses, strict=True):
            fragment = by_id[relation["source_artifact_id"]]
            fragment_payload = _read_bytes_with_long_path(staging_root / Path(fragment["locator"]))
            if status != "success" and fragment_payload:
                raise RuntimeError(f"packaged_physical_page_empty_placeholder_invalid:{status}:{fragment}")
            if status == "success" and not fragment_payload:
                raise RuntimeError(f"packaged_physical_page_success_fragment_empty:{fragment}")
            if fragment_payload and fragment_payload in primary_payload:
                raise RuntimeError("packaged_physical_page_primary_duplicates_ocr_fragment")
    resolved_resource_pages: list[int] = []
    unresolved_resource_ids: list[str] = []
    for relation in resource_relations:
        semantics = relation.get("page_resource")
        if semantics is None:
            if relation.get("target_artifact_id") != primary_id:
                raise RuntimeError(f"packaged_physical_page_unresolved_owner_invalid:{relation}")
            unresolved_resource_ids.append(relation["source_artifact_id"])
            continue
        if (
            not isinstance(semantics, dict)
            or set(semantics) != {"source_page"}
            or type(semantics.get("source_page")) is not int
        ):
            raise RuntimeError(f"packaged_physical_page_resource_semantics_invalid:{relation}")
        resolved_resource_pages.append(semantics["source_page"])
        expected_owner = next(
            (
                item["source_artifact_id"]
                for item in page_relations
                if item["page_fragment"]["source_page"] == semantics["source_page"]
            ),
            primary_id,
        )
        if relation.get("target_artifact_id") != expected_owner:
            raise RuntimeError(f"packaged_physical_page_resource_owner_invalid:{relation}")
    if any(page < 1 or page > page_count for page in resolved_resource_pages):
        raise RuntimeError(f"packaged_physical_page_resource_range_invalid:{resolved_resource_pages}")
    if keep_images and page_count == 4 and sorted(resolved_resource_pages) != [1, 2, 3, 4]:
        raise RuntimeError(f"packaged_physical_page_resource_coverage_invalid:{resolved_resource_pages}")
    unresolved_diagnostics = [
        diagnostic
        for diagnostic in diagnostics
        if isinstance(diagnostic, dict) and diagnostic.get("code") == "resource_page_unresolved"
    ]
    unresolved_diagnostic_ids = [
        artifact_id
        for diagnostic in unresolved_diagnostics
        for artifact_id in (diagnostic.get("artifact_id"),)
        if isinstance(artifact_id, str)
    ]
    if any(
        diagnostic.get("artifact_id") not in by_id
        for diagnostic in diagnostics
        if isinstance(diagnostic, dict) and "artifact_id" in diagnostic
    ):
        raise RuntimeError(f"packaged_physical_page_diagnostic_artifact_invalid:{diagnostics}")
    if sorted(unresolved_resource_ids) != sorted(unresolved_diagnostic_ids):
        raise RuntimeError(
            f"packaged_physical_page_unresolved_diagnostic_invalid:{unresolved_resource_ids}:{unresolved_diagnostics}"
        )
    if ocr_enabled:
        expected_ocr_diagnostic_ids = sorted(relation["source_artifact_id"] for relation in page_relations)
        actual_ocr_diagnostic_ids = sorted(
            artifact_id
            for diagnostic in diagnostics
            if isinstance(diagnostic, dict) and diagnostic.get("code") == "OCR-BEST-EFFORT"
            for artifact_id in (diagnostic.get("artifact_id"),)
            if isinstance(artifact_id, str)
        )
        if actual_ocr_diagnostic_ids != expected_ocr_diagnostic_ids:
            raise RuntimeError(
                "packaged_physical_page_ocr_diagnostic_invalid:"
                f"{expected_ocr_diagnostic_ids}:{actual_ocr_diagnostic_ids}"
            )
    elif any(
        isinstance(diagnostic, dict) and diagnostic.get("code") == "OCR-BEST-EFFORT" for diagnostic in diagnostics
    ):
        raise RuntimeError(f"packaged_physical_page_unexpected_ocr_diagnostic:{diagnostics}")
    for artifact_id, artifact in by_id.items():
        locator = artifact.get("locator")
        locator_parts = locator.split("/") if isinstance(locator, str) else []
        if (
            not isinstance(locator, str)
            or not locator
            or locator.startswith("/")
            or "\\" in locator
            or ":" in locator_parts[0]
            or any(part in {"", ".", ".."} for part in locator_parts)
        ):
            raise RuntimeError(f"packaged_physical_page_locator_invalid:{artifact}")
        output_path = staging_root / Path(locator)
        payload = _read_bytes_with_long_path(output_path)
        if len(payload) != artifact.get("size_bytes") or hashlib.sha256(payload).hexdigest() != artifact.get("sha256"):
            raise RuntimeError(f"packaged_physical_page_integrity_invalid:{artifact_id}")
    return bundle


def _native_long_path(path: Path) -> str:
    """Return an absolute native path that bypasses legacy Win32 MAX_PATH."""
    native = os.path.abspath(os.fspath(path))
    if sys.platform == "win32" and not native.startswith("\\\\?\\"):
        native = f"\\\\?\\UNC\\{native[2:]}" if native.startswith("\\\\") else f"\\\\?\\{native}"
    return native


def _read_bytes_with_long_path(path: Path) -> bytes:
    """Read verifier artifacts beyond legacy Win32 MAX_PATH without changing wire locators."""

    native = _native_long_path(path)
    with open(native, "rb") as stream:
        return stream.read()


def _zipfile_with_long_path(path: Path) -> zipfile.ZipFile:
    """Open an existing ZIP/OOXML payload without reopening its Win32 path."""

    return zipfile.ZipFile(io.BytesIO(_read_bytes_with_long_path(path)))


def _read_text_with_long_path(path: Path) -> str:
    text = _read_bytes_with_long_path(path).decode("utf-8")
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _find_markdown_outputs_with_long_path(root: Path) -> list[Path]:
    outputs: list[Path] = []
    for directory, _directories, filenames in os.walk(_native_long_path(root)):
        outputs.extend(Path(directory) / filename for filename in filenames if filename.lower().endswith(".md"))
    return outputs


def _remove_tree_with_long_path(path: Path) -> None:
    shutil.rmtree(_native_long_path(path), ignore_errors=True)


@contextlib.contextmanager
def _temporary_directory_with_long_path_cleanup(*, prefix: str) -> Any:
    path = Path(tempfile.mkdtemp(prefix=prefix))
    try:
        yield os.fspath(path)
    finally:
        _remove_tree_with_long_path(path)


def _write_ocr_png(path: Path) -> None:
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (900, 240), "white")
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 56)
    except OSError:
        font = ImageFont.load_default()
    draw.text((36, 72), "HELLO DOCWEN OCR", fill="black", font=font)
    image.save(path)


def _write_proofread_report_fixture(path: Path) -> bytes:
    """Write the byte-stable Markdown fixture used by the packaged report gate."""
    raw = _PROOFREAD_REPORT_FIXTURE_TEXT.encode("utf-8")
    path.write_bytes(raw)
    return raw


def _run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["DOCWEN_CONFIG_DIR"] = str(cwd / "config_home")
    env["DOCWEN_LOG_DIR"] = str(cwd / "log_home")
    env["DOCWEN_LOG_TO_TEMP"] = ""
    return subprocess.run(
        [str(binary_path), *args],
        cwd=cwd,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )


def _run_multiprocessing_egress_boundary_smoke(binary_path: Path, *, work_dir: Path) -> Path:
    """Prove the frozen parent is guarded while spawned helpers are not."""

    # Keep this verifier-only report comfortably below the legacy Win32
    # MAX_PATH boundary.  ``work_dir`` is intentionally already a 200+ character
    # Unicode path; a descriptive report filename must not accidentally turn
    # the multiprocessing check into a separate OS long-path-policy check.
    report_path = work_dir / "egress.json"
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["DOCWEN_CONFIG_DIR"] = str(work_dir / "config_home")
    env["DOCWEN_LOG_DIR"] = str(work_dir / "log_home")
    env["DOCWEN_LOG_TO_TEMP"] = ""
    env["DOCWEN_TEST_MULTIPROCESS_EGRESS_REPORT"] = str(report_path)
    completed = subprocess.run(
        [str(binary_path)],
        cwd=work_dir,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        timeout=120,
    )
    _verify_process_succeeded(completed, command_name="multiprocessing egress boundary")
    if not report_path.is_file():
        raise RuntimeError("packaged_multiprocessing_egress_report_missing")
    payload = json.loads(_read_text_with_long_path(report_path))
    if not isinstance(payload, dict):
        raise RuntimeError(f"packaged_multiprocessing_egress_boundary_invalid:{payload}")
    parent_guard = payload.get("parent_guard")
    child = payload.get("child")
    child_guard = child.get("guard") if isinstance(child, dict) else None
    child_audit_probe_allowed = child.get("audit_probe_allowed") if isinstance(child, dict) else None
    if (
        not isinstance(parent_guard, dict)
        or parent_guard.get("state") != "enforced"
        or parent_guard.get("bootstrap") != "pyinstaller_runtime_hook"
        or payload.get("parent_audit_probe_blocked") is not True
        or payload.get("child_exit_code") != 0
        or not isinstance(child_guard, dict)
        or child_guard.get("state") != "not_installed"
        or child_guard.get("bootstrap") != "none"
        or child_audit_probe_allowed is not True
    ):
        raise RuntimeError(f"packaged_multiprocessing_egress_boundary_invalid:{payload}")
    return report_path


def _load_json_payload(proc: subprocess.CompletedProcess[str], *, command_name: str) -> dict[str, object]:
    _verify_process_succeeded(proc, command_name=command_name)
    return _decode_json_payload(proc, command_name=command_name)


def _decode_json_payload(proc: subprocess.CompletedProcess[str], *, command_name: str) -> dict[str, object]:
    try:
        payload = json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"{command_name} did not emit valid JSON: {proc.stdout}\n{proc.stderr}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError(f"{command_name} did not emit a JSON object: {proc.stdout}\n{proc.stderr}")
    return payload


def _load_json_failure_payload(
    proc: subprocess.CompletedProcess[str],
    *,
    command_name: str,
) -> dict[str, object]:
    """Require one typed machine failure from a packaged process."""
    _verify_no_plugin_load_failure(proc, command_name=command_name)
    if proc.returncode == 0:
        raise RuntimeError(f"{command_name} unexpectedly succeeded: {proc.stdout}\n{proc.stderr}")
    if proc.stderr:
        raise RuntimeError(f"{command_name} machine failure wrote stderr: {proc.stderr}")

    payload = _decode_json_payload(proc, command_name=command_name)
    if payload.get("success") is not False:
        raise RuntimeError(f"{command_name} failure payload did not report success=false: {payload}")
    error = payload.get("error")
    if not isinstance(error, dict):
        raise RuntimeError(f"{command_name} failure payload missing error object: {payload}")
    if not isinstance(error.get("category"), str) or not error["category"]:
        raise RuntimeError(f"{command_name} failure payload missing error category: {payload}")
    if not isinstance(error.get("code"), str) or not error["code"]:
        raise RuntimeError(f"{command_name} failure payload missing error code: {payload}")
    return payload


def _verify_process_succeeded(proc: subprocess.CompletedProcess[str], *, command_name: str) -> None:
    """Fail closed on packaged process or dynamic-plugin loading failures."""
    _verify_no_plugin_load_failure(proc, command_name=command_name)
    if proc.returncode != 0:
        raise RuntimeError(f"{command_name} failed with exit code {proc.returncode}: {proc.stdout}\n{proc.stderr}")


def _verify_no_plugin_load_failure(proc: subprocess.CompletedProcess[str], *, command_name: str) -> None:
    combined_output = f"{proc.stdout}\n{proc.stderr}"
    if _PLUGIN_LOAD_FAILURE_MARKER in combined_output:
        failures = [line.strip() for line in combined_output.splitlines() if _PLUGIN_LOAD_FAILURE_MARKER in line]
        raise RuntimeError(f"{command_name} reported unavailable packaged plugins: {'; '.join(failures)}")


def _resolve_output_file(payload: dict[str, object], *, work_dir: Path, command_name: str) -> Path:
    data = payload.get("data")
    if not isinstance(data, dict):
        raise RuntimeError(f"{command_name} payload missing data object: {payload}")

    output_file_raw = data.get("output")
    if not isinstance(output_file_raw, str):
        raise RuntimeError(f"{command_name} payload missing output: {payload}")

    output_file = Path(output_file_raw)
    if not output_file.is_absolute():
        output_file = (work_dir / output_file).resolve()
    return output_file


def _verify_dependency_egress_guard(data: dict[str, object], *, command_name: str) -> None:
    """Require the packaged process to report the exact enforced egress boundary."""

    security = data.get("security")
    guard = security.get("dependency_egress_guard") if isinstance(security, dict) else None
    if not isinstance(guard, dict):
        raise RuntimeError(f"{command_name} omitted dependency egress guard status")
    expected = {
        "state": "enforced",
        "installed": True,
        "active": True,
        "scope": "docwen_python_process",
        "policy": "deny_dns_and_ip",
        "mechanism": "cpython_audit_hook",
        "bootstrap": "pyinstaller_runtime_hook",
        "external_processes": "not_managed",
    }
    mismatches = {key: guard.get(key) for key, value in expected.items() if guard.get(key) != value}
    transports = guard.get("local_transports")
    if (
        mismatches
        or not isinstance(transports, list)
        or set(transports)
        != {
            "windows_named_pipe",
            "unix_domain_socket",
        }
    ):
        raise RuntimeError(f"{command_name} dependency egress guard is not enforced: {guard}")


def _verify_capability_discovery(binary_path: Path, *, work_dir: Path) -> None:
    payload = _load_json_payload(
        _run(binary_path, "resources", "list", "formats", "--json", "--quiet", cwd=work_dir),
        command_name="resources list formats",
    )
    if payload.get("success") is not True or payload.get("command") != "resources list":
        raise RuntimeError(f"resources list formats returned an invalid success envelope: {payload}")
    data = payload.get("data")
    if not isinstance(data, dict):
        raise RuntimeError(f"resources list formats payload missing data object: {payload}")
    if data.get("contract") != {"id": "docwen.runtime-capabilities", "version": 1}:
        raise RuntimeError(f"resources list formats capability contract mismatch: {data.get('contract')}")
    runtime = data.get("runtime")
    if not isinstance(runtime, dict) or runtime.get("state") != "available":
        raise RuntimeError(f"resources list formats runtime state mismatch: {runtime}")
    _verify_dependency_egress_guard(data, command_name="resources list formats")
    sources = data.get("sources")
    if not isinstance(sources, list) or not sources:
        raise RuntimeError("resources list formats did not expose the packaged runtime composition")
    routes = [
        route
        for source in sources
        if isinstance(source, dict)
        for route in source.get("routes", [])
        if isinstance(route, dict)
    ]
    if not routes:
        raise RuntimeError("resources list formats did not expose any packaged routes")
    required_route_fields = {
        "id",
        "operation",
        "source",
        "target",
        "action",
        "plugin",
        "available",
        "state",
        "platforms",
        "platform_supported",
        "required_capabilities",
        "optional_capabilities",
        "limitations",
        "options",
    }
    if any(not required_route_fields.issubset(route) for route in routes):
        raise RuntimeError("resources list formats route contract is incomplete")
    available_routes = [route for route in routes if route.get("available") is True]
    if not available_routes:
        raise RuntimeError("resources list formats did not expose any available packaged routes")
    if not any(
        route.get("operation") == "conversion"
        and route.get("source") == "pdf"
        and route.get("target") == "md"
        and route.get("action") in (None, "")
        and route.get("available") is True
        for route in routes
    ):
        raise RuntimeError("resources list formats did not expose an available PDF to Markdown route")
    if not any(route.get("operation") == "action" and route.get("action") for route in routes):
        raise RuntimeError("resources list formats omitted action-only routes")
    gates = data.get("gates")
    if not isinstance(gates, list):
        raise RuntimeError("resources list formats capability payload omitted dependency gates")
    pymupdf_layout_gate = next(
        (gate for gate in gates if isinstance(gate, dict) and gate.get("id") == _PYMUPDF_LAYOUT_GATE_ID),
        None,
    )
    if not isinstance(pymupdf_layout_gate, dict) or pymupdf_layout_gate.get("available") is not True:
        raise RuntimeError("resources list formats reported python.pymupdf4llm unavailable")
    counts = data.get("counts")
    if (
        not isinstance(counts, dict)
        or counts.get("routes") != len(routes)
        or counts.get("available_routes") != len(available_routes)
    ):
        raise RuntimeError(f"resources list formats count mismatch: {counts}")


def _verify_optimization_discovery(binary_path: Path, *, work_dir: Path) -> None:
    payload = _load_json_payload(
        _run(binary_path, "resources", "list", "optimizations", "--json", "--quiet", cwd=work_dir),
        command_name="resources list optimizations",
    )
    if payload.get("success") is not True or payload.get("command") != "resources list":
        raise RuntimeError(f"resources list optimizations returned an invalid success envelope: {payload}")
    data = payload.get("data")
    if not isinstance(data, dict):
        raise RuntimeError(f"resources list optimizations payload missing data object: {payload}")
    if data.get("resource") != "optimizations":
        raise RuntimeError("resources list optimizations resource discriminator mismatch")
    if data.get("contract") != {"id": "docwen.optimizations", "version": 1}:
        raise RuntimeError(f"resources list optimizations contract mismatch: {data.get('contract')}")
    runtime = data.get("runtime")
    if not isinstance(runtime, dict) or runtime.get("state") != "available":
        raise RuntimeError(f"resources list optimizations runtime state mismatch: {runtime}")

    resources = data.get("resources")
    if not isinstance(resources, list) or not resources:
        raise RuntimeError("resources list optimizations did not expose the packaged optimizer composition")
    resource_fields = {"id", "name", "action_name", "scopes", "available", "state", "bindings"}
    binding_fields = {
        "scope",
        "route_id",
        "source",
        "source_category",
        "target",
        "available",
        "state",
    }
    if any(not isinstance(resource, dict) or set(resource) != resource_fields for resource in resources):
        raise RuntimeError("resources list optimizations resource contract is incomplete")
    if any(
        not isinstance(resource["bindings"], list)
        or not resource["bindings"]
        or any(not isinstance(binding, dict) for binding in resource["bindings"])
        for resource in resources
    ):
        raise RuntimeError("resources list optimizations binding contract is incomplete")
    bindings = [binding for resource in resources for binding in resource["bindings"]]
    if any(set(binding) != binding_fields for binding in bindings):
        raise RuntimeError("resources list optimizations binding contract is incomplete")

    by_id = {str(resource["id"]): resource for resource in resources}
    expected = {
        "gongwen": {
            "action_name": "gongwen",
            "scopes": {"document_to_md"},
            "bindings": {("docx", "document", "md")},
        },
        "invoice_cn": {
            "action_name": "invoice_cn",
            "scopes": {"layout_to_md", "image_to_md"},
            "bindings": {
                ("pdf", "layout", "md"),
                ("ofd", "layout", "md"),
                ("image", "image", "md"),
            },
        },
    }
    for resource_id, expected_resource in expected.items():
        resource = by_id.get(resource_id)
        if resource is None:
            raise RuntimeError(f"resources list optimizations omitted packaged resource: {resource_id}")
        actual_bindings = {
            (binding.get("source"), binding.get("source_category"), binding.get("target"))
            for binding in resource["bindings"]
            if isinstance(binding, dict)
        }
        if (
            resource.get("action_name") != expected_resource["action_name"]
            or set(resource.get("scopes", [])) != expected_resource["scopes"]
            or actual_bindings != expected_resource["bindings"]
        ):
            raise RuntimeError(f"resources list optimizations binding mismatch for {resource_id}: {resource}")

    counts = data.get("counts")
    available_resources = sum(resource.get("available") is True for resource in resources)
    available_bindings = sum(binding.get("available") is True for binding in bindings)
    expected_counts = {
        "resources": len(resources),
        "available_resources": available_resources,
        "unavailable_resources": len(resources) - available_resources,
        "bindings": len(bindings),
        "available_bindings": available_bindings,
        "unavailable_bindings": len(bindings) - available_bindings,
    }
    if counts != expected_counts:
        raise RuntimeError(f"resources list optimizations count mismatch: {counts}")


def _run_template_resource_smoke(binary_path: Path, *, work_dir: Path) -> Path:
    """Verify packaged template discovery and canonical-ID conversion end to end."""

    payload = _load_json_payload(
        _run(binary_path, "resources", "list", "templates", "--json", "--quiet", cwd=work_dir),
        command_name="resources list templates",
    )
    if payload.get("success") is not True or payload.get("command") != "resources list":
        raise RuntimeError(f"resources list templates returned an invalid success envelope: {payload}")
    data = payload.get("data")
    if not isinstance(data, dict) or set(data) != {"type", "resources", "total"}:
        raise RuntimeError(f"resources list templates payload contract mismatch: {data}")
    if data.get("type") != "templates":
        raise RuntimeError(f"resources list templates type mismatch: {data.get('type')}")
    resources = data.get("resources")
    total = data.get("total")
    if not isinstance(resources, list) or not resources:
        raise RuntimeError("resources list templates did not expose packaged templates")
    if type(total) is not int or total != len(resources) or total != len(_REQUIRED_TEMPLATE_FILES):
        raise RuntimeError(f"resources list templates count mismatch: total={total}, resources={len(resources)}")

    identifiers: list[str] = []
    docx_resources: list[dict[str, object]] = []
    for resource in resources:
        if not isinstance(resource, dict) or set(resource) != _TEMPLATE_RESOURCE_FIELDS:
            raise RuntimeError(f"resources list templates resource contract is incomplete: {resource}")
        template_id = resource.get("id")
        name = resource.get("name")
        target = resource.get("target")
        description = resource.get("description")
        resource_path = resource.get("path")
        size_bytes = resource.get("size_bytes")
        modified_ns = resource.get("modified_ns")
        if (
            not isinstance(template_id, str)
            or _TEMPLATE_ID_PATTERN.fullmatch(template_id) is None
            or not isinstance(name, str)
            or not name
            or target not in {"docx", "xlsx"}
            or not template_id.startswith(f"template.{target}.")
            or not isinstance(description, str)
            or not isinstance(resource_path, str)
            or not resource_path
            or type(size_bytes) is not int
            or size_bytes < 0
            or type(modified_ns) is not int
            or modified_ns < 0
        ):
            raise RuntimeError(f"resources list templates resource value is invalid: {resource}")
        identifiers.append(template_id)
        if target == "docx":
            docx_resources.append(resource)

    if len(set(identifiers)) != len(identifiers):
        raise RuntimeError("resources list templates returned duplicate canonical IDs")
    expected_targets = {
        Path(filename).stem: Path(filename).suffix.removeprefix(".").casefold() for filename in _REQUIRED_TEMPLATE_FILES
    }
    actual_targets = {str(resource["name"]): str(resource["target"]) for resource in resources}
    if actual_targets != expected_targets:
        raise RuntimeError(
            "resources list templates did not match the packaged template manifest: "
            f"expected={expected_targets}, actual={actual_targets}"
        )
    if not docx_resources:
        raise RuntimeError("resources list templates did not expose a DOCX template")

    selected = min(docx_resources, key=lambda resource: str(resource["id"]))
    template_id = str(selected["id"])
    show_payload = _load_json_payload(
        _run(
            binary_path,
            "resources",
            "show",
            "templates",
            template_id,
            "--json",
            "--quiet",
            cwd=work_dir,
        ),
        command_name="resources show templates",
    )
    if show_payload.get("success") is not True or show_payload.get("command") != "resources show":
        raise RuntimeError(f"resources show templates returned an invalid success envelope: {show_payload}")
    show_data = show_payload.get("data")
    if not isinstance(show_data, dict) or set(show_data) != {"type", "resource"}:
        raise RuntimeError(f"resources show templates payload contract mismatch: {show_data}")
    if show_data.get("type") != "templates":
        raise RuntimeError(f"resources show templates type mismatch: {show_data.get('type')}")
    shown_resource = show_data.get("resource")
    if not isinstance(shown_resource, dict) or set(shown_resource) != _TEMPLATE_RESOURCE_FIELDS:
        raise RuntimeError(f"resources show templates resource contract is incomplete: {shown_resource}")
    if shown_resource != selected:
        raise RuntimeError(
            "resources show templates did not exactly match the listed resource: "
            f"listed={selected}, shown={shown_resource}"
        )

    source_path = work_dir / "canonical-template-id-smoke.md"
    source_path.write_text(f"# {_TEMPLATE_SMOKE_TEXT}\n\nTemplate resource ID closed loop.\n", encoding="utf-8")
    output_path = work_dir / "canonical-template-id-smoke.docx"
    convert_payload = _load_json_payload(
        _run(
            binary_path,
            "convert",
            str(source_path),
            "--to",
            "docx",
            "--template",
            template_id,
            "--output",
            str(output_path),
            "--json",
            "--quiet",
            cwd=work_dir,
        ),
        command_name="convert with canonical template ID",
    )
    if convert_payload.get("success") is not True or convert_payload.get("command") != "convert":
        raise RuntimeError(f"convert with canonical template ID failed: {convert_payload}")
    resolved_output = _resolve_output_file(
        convert_payload,
        work_dir=work_dir,
        command_name="convert with canonical template ID",
    )
    if resolved_output.resolve() != output_path.resolve():
        raise RuntimeError(
            "convert with canonical template ID returned an unexpected output: "
            f"expected={output_path.resolve()}, actual={resolved_output.resolve()}"
        )
    try:
        with _zipfile_with_long_path(resolved_output) as package:
            document_xml = package.read("word/document.xml").decode("utf-8", errors="replace")
    except (OSError, KeyError, zipfile.BadZipFile) as exc:
        raise RuntimeError(
            f"convert with canonical template ID did not create a valid DOCX: {resolved_output}"
        ) from exc
    if _TEMPLATE_SMOKE_TEXT not in document_xml:
        raise RuntimeError("convert with canonical template ID output omitted the smoke fixture text")
    return resolved_output


def _verify_doctor_payload(payload: dict[str, object]) -> None:
    """Require healthy base probes and one canonical packaged runtime projection.

    Optional features and host-provided Office backends are represented by the
    projection and must not make an otherwise healthy package fail this gate.
    Feature-specific release requirements are checked explicitly below.
    """
    if payload.get("success") is not True or payload.get("command") != "doctor":
        raise RuntimeError(f"doctor returned an invalid success envelope: {payload}")
    data = payload.get("data")
    if not isinstance(data, dict):
        raise RuntimeError(f"doctor payload missing data object: {payload}")
    if data.get("all_ok") is not True:
        raise RuntimeError(f"doctor did not report all_ok=true: {payload}")

    checks = data.get("checks")
    if not isinstance(checks, list):
        raise RuntimeError("doctor payload omitted diagnostic checks")
    checks_by_id: dict[str, object] = {}
    for check in checks:
        if not isinstance(check, dict):
            continue
        check_id = check.get("id")
        if isinstance(check_id, str):
            checks_by_id[check_id] = check
    if len(checks) != len(_DOCTOR_BASE_CHECK_IDS) or set(checks_by_id) != _DOCTOR_BASE_CHECK_IDS:
        raise RuntimeError(f"doctor base check set mismatch: {sorted(checks_by_id)}")
    for check_id in sorted(_DOCTOR_BASE_CHECK_IDS):
        check = checks_by_id[check_id]
        if not isinstance(check, dict) or check.get("status") != "ok":
            raise RuntimeError(f"doctor base check unavailable: {check_id}")

    capability_summary = data.get("capability_summary")
    if not isinstance(capability_summary, dict):
        raise RuntimeError("doctor capability summary is not an object")
    _verify_dependency_egress_guard(capability_summary, command_name="doctor")
    gates = capability_summary.get("gates")
    if not isinstance(gates, list):
        raise RuntimeError("doctor capability summary omitted dependency gates")
    layout_gate = next(
        (gate for gate in gates if isinstance(gate, dict) and gate.get("id") == _PYMUPDF_LAYOUT_GATE_ID),
        None,
    )
    if not isinstance(layout_gate, dict) or layout_gate.get("available") is not True:
        raise RuntimeError("doctor capability summary reported python.pymupdf4llm unavailable")


def _load_verified_doctor_payload(proc: subprocess.CompletedProcess[str]) -> dict[str, object]:
    """Decode doctor JSON before interpreting its process-level health status."""
    _verify_no_plugin_load_failure(proc, command_name="doctor")
    payload = _decode_json_payload(proc, command_name="doctor")
    _verify_doctor_payload(payload)
    if proc.returncode != 0:
        raise RuntimeError(
            f"doctor reported a healthy payload but exited with {proc.returncode}: {proc.stdout}\n{proc.stderr}"
        )
    return payload


def _verify_md_output_file(payload: dict[str, object], *, work_dir: Path, command_name: str) -> Path:
    output_file = _resolve_output_file(payload, work_dir=work_dir, command_name=command_name)
    if not os.path.isfile(_native_long_path(output_file)):
        raise RuntimeError(f"packaged_convert_output_missing: {output_file}")
    if output_file.suffix.lower() != ".md":
        raise RuntimeError(f"packaged_convert_output_suffix_unexpected: {output_file}")
    if not _read_text_with_long_path(output_file).strip():
        raise RuntimeError(f"packaged_convert_output_empty: {output_file}")
    return output_file


def _run_pymupdf_layout_smoke(binary_path: Path, *, work_dir: Path) -> Path:
    """Trigger the packaged PDF-to-Markdown route and its lazy-loaded layout model."""
    source = work_dir / "PyMuPDF Layout 最小验证.pdf"
    output = work_dir / "PyMuPDF Layout 最小验证.md"
    _write_pymupdf_layout_pdf(source)

    payload = _load_json_payload(
        _run(
            binary_path,
            "convert",
            str(source),
            "--to",
            "md",
            "--output",
            str(output),
            "--json",
            "--quiet",
            cwd=work_dir,
        ),
        command_name="convert PDF with PyMuPDF Layout",
    )
    if payload.get("success") is not True or payload.get("command") != "convert":
        raise RuntimeError(f"convert PDF with PyMuPDF Layout returned an invalid success envelope: {payload}")
    output_file = _verify_md_output_file(
        payload,
        work_dir=work_dir,
        command_name="convert PDF with PyMuPDF Layout",
    )
    if _PYMUPDF_LAYOUT_SMOKE_TEXT not in _read_text_with_long_path(output_file):
        raise RuntimeError(f"packaged_pymupdf_layout_output_missing_expected_text: {output_file}")
    return output_file


def _inspect_content_contract(
    binary_path: Path,
    *,
    work_dir: Path,
    input_path: Path,
    expected: dict[str, str],
    command_name: str,
) -> dict[str, object]:
    payload = _load_json_payload(
        _run(binary_path, "inspect", str(input_path), "--json", "--quiet", cwd=work_dir),
        command_name=command_name,
    )
    if payload.get("success") is not True or payload.get("command") != "inspect":
        raise RuntimeError(f"{command_name} returned an invalid success envelope: {payload}")
    data = payload.get("data")
    if not isinstance(data, dict):
        raise RuntimeError(f"{command_name} payload missing data object: {payload}")
    for field, expected_value in expected.items():
        if data.get(field) != expected_value:
            raise RuntimeError(f"{command_name} {field} mismatch: expected {expected_value!r}, got {data.get(field)!r}")
    return payload


def _verify_blocked_container_failure(
    binary_path: Path,
    *,
    work_dir: Path,
    input_path: Path,
    output_path: Path,
    command_name: str,
) -> None:
    proc = _run(
        binary_path,
        "convert",
        str(input_path),
        "--to",
        "md",
        "--output",
        str(output_path),
        "--use-detected-format",
        "--json",
        "--quiet",
        cwd=work_dir,
    )
    if proc.returncode != 2:
        raise RuntimeError(f"{command_name} expected exit 2, got {proc.returncode}: {proc.stdout}\n{proc.stderr}")
    payload = _load_json_failure_payload(proc, command_name=command_name)
    if payload.get("command") != "convert":
        raise RuntimeError(f"{command_name} failure payload command mismatch: {payload}")
    error = payload["error"]
    if not isinstance(error, dict):  # Narrowed by _load_json_failure_payload; retained for static analyzers.
        raise RuntimeError(f"{command_name} failure payload missing error object: {payload}")
    if error.get("category") != "invalid_input" or error.get("code") != "file_container_invalid":
        raise RuntimeError(f"{command_name} returned the wrong typed failure: {payload}")
    details = error.get("details")
    admission = details.get("admission") if isinstance(details, dict) else None
    if not isinstance(admission, dict):
        raise RuntimeError(f"{command_name} failure payload missing admission details: {payload}")
    if admission.get("decision") != "block" or admission.get("reason_code") != "FILE_CONTAINER_INVALID":
        raise RuntimeError(f"{command_name} failure payload contains the wrong admission decision: {payload}")
    if output_path.exists():
        raise RuntimeError(f"{command_name} published output despite blocked input: {output_path}")


def _run_content_first_contract_smoke(binary_path: Path, *, work_dir: Path) -> Path:
    """Prove the installed CLI follows content, not a filename suffix."""
    disguised_xlsx = work_dir / "实际为 XLSX 的文本后缀.txt"
    disguised_output = work_dir / "伪装表格转换结果.md"
    _write_xlsx(disguised_xlsx)
    _inspect_content_contract(
        binary_path,
        work_dir=work_dir,
        input_path=disguised_xlsx,
        expected={
            "declared_format": "txt",
            "detected_format": "xlsx",
            "relation": "cross_family_mismatch",
            "decision": "require_explicit_acceptance",
        },
        command_name="inspect disguised XLSX",
    )
    convert_payload = _load_json_payload(
        _run(
            binary_path,
            "convert",
            str(disguised_xlsx),
            "--to",
            "md",
            "--output",
            str(disguised_output),
            "--use-detected-format",
            "--json",
            "--quiet",
            cwd=work_dir,
        ),
        command_name="convert disguised XLSX",
    )
    if convert_payload.get("success") is not True or convert_payload.get("command") != "convert":
        raise RuntimeError(f"convert disguised XLSX returned an invalid success envelope: {convert_payload}")
    converted_file = _verify_md_output_file(
        convert_payload,
        work_dir=work_dir,
        command_name="convert disguised XLSX",
    )
    converted_text = _read_text_with_long_path(converted_file)
    if not all(value in converted_text for value in ("name", "value", "alpha")):
        raise RuntimeError(f"packaged_disguised_xlsx_output_missing_expected_table_content: {converted_file}")

    plain_text = "Plain text without Markdown markers.\nSecond plain line.\n"
    markdown_text = "# Heading\n\n**bold** text\n"
    text_fixtures = (
        (
            work_dir / "纯文本内容.txt",
            plain_text,
            {
                "declared_format": "txt",
                "detected_format": "txt",
                "relation": "exact_match",
                "decision": "allow",
            },
        ),
        (
            work_dir / "纯文本内容.md",
            plain_text,
            {
                "declared_format": "markdown",
                "detected_format": "txt",
                "relation": "compatible_text",
                "decision": "allow_with_warning",
            },
        ),
        (
            work_dir / "Markdown 内容.txt",
            markdown_text,
            {
                "declared_format": "txt",
                "detected_format": "markdown",
                "relation": "compatible_text",
                "decision": "allow_with_warning",
            },
        ),
        (
            work_dir / "Markdown 内容.md",
            markdown_text,
            {
                "declared_format": "markdown",
                "detected_format": "markdown",
                "relation": "equivalent_alias",
                "decision": "allow",
            },
        ),
    )
    for input_path, content, expected in text_fixtures:
        input_path.write_text(content, encoding="utf-8")
        _inspect_content_contract(
            binary_path,
            work_dir=work_dir,
            input_path=input_path,
            expected=expected,
            command_name=f"inspect text fixture {input_path.name}",
        )

    ordinary_zip = work_dir / "普通 ZIP 伪装文档.docx"
    with zipfile.ZipFile(ordinary_zip, "w") as package:
        package.writestr("hello.txt", "hello")
    corrupt_ooxml = work_dir / "损坏 OOXML 伪装表格.xlsx"
    corrupt_ooxml.write_bytes(b"PK\x03\x04not-a-valid-central-directory")
    _verify_blocked_container_failure(
        binary_path,
        work_dir=work_dir,
        input_path=ordinary_zip,
        output_path=work_dir / "普通 ZIP 不应生成.md",
        command_name="convert ordinary ZIP disguised as DOCX",
    )
    _verify_blocked_container_failure(
        binary_path,
        work_dir=work_dir,
        input_path=corrupt_ooxml,
        output_path=work_dir / "损坏 OOXML 不应生成.md",
        command_name="convert corrupt OOXML disguised as XLSX",
    )
    return converted_file


def _run_optional_ocr_smoke(binary_path: Path, *, work_dir: Path) -> Path:
    source = work_dir / "sample_ocr.png"
    output = work_dir / "sample_ocr.md"
    _write_ocr_png(source)
    payload = _load_json_payload(
        _run(
            binary_path,
            "convert",
            str(source),
            "--to",
            "md",
            "--output",
            str(output),
            "--ocr",
            "--ocr-placement",
            "main_md",
            "--image-mode",
            "file",
            "--json",
            "--quiet",
            cwd=work_dir,
        ),
        command_name="ocr convert",
    )
    if payload.get("success") is not True:
        raise RuntimeError(f"ocr convert returned success=false: {payload}")
    output_file = _verify_md_output_file(payload, work_dir=work_dir, command_name="ocr convert")
    content = _read_text_with_long_path(output_file)
    if "HELLO" not in content.upper() or "OCR" not in content.upper():
        raise RuntimeError(f"packaged_ocr_output_missing_expected_text: {output_file}")
    return output_file


def _run_numbering_smoke(binary_path: Path, *, work_dir: Path) -> tuple[Path, Path]:
    del work_dir
    # The v0.9 input contract intentionally rejects user-supplied paths above
    # 259 UTF-16 units. Keep this round-trip fixture short; long output paths
    # are exercised independently by the conversion and Machine gates.
    with _temporary_directory_with_long_path_cleanup(prefix="dw-numbering-") as short_dir:
        return _run_numbering_smoke_impl(binary_path, work_dir=Path(short_dir))


def _run_numbering_smoke_impl(binary_path: Path, *, work_dir: Path) -> tuple[Path, Path]:
    source = work_dir / "编号 输入.md"
    added = work_dir / "编号 添加.md"
    removed = work_dir / "编号 移除.md"
    original = "# 已有标题\n\n## 二级标题\n\n正文。\n"
    source.write_text(original, encoding="utf-8")

    add_payload = _load_json_payload(
        _run(
            binary_path,
            "number",
            "markdown",
            str(source),
            "--operation",
            "add",
            "--scheme",
            "gongwen_standard",
            "--output",
            str(added),
            "--json",
            "--quiet",
            cwd=work_dir,
        ),
        command_name="number markdown add",
    )
    if add_payload.get("success") is not True:
        raise RuntimeError(f"packaged_number_add_failed: {add_payload}")
    added_output = _verify_md_output_file(add_payload, work_dir=work_dir, command_name="number markdown add")
    added_text = _read_text_with_long_path(added_output)
    if "# 一、已有标题" not in added_text or "## （一）二级标题" not in added_text:
        raise RuntimeError(f"packaged_number_add_content_mismatch: {added_text!r}")

    remove_payload = _load_json_payload(
        _run(
            binary_path,
            "number",
            "markdown",
            str(added_output),
            "--operation",
            "remove",
            "--output",
            str(removed),
            "--json",
            "--quiet",
            cwd=work_dir,
        ),
        command_name="number markdown remove",
    )
    if remove_payload.get("success") is not True:
        raise RuntimeError(f"packaged_number_remove_failed: {remove_payload}")
    removed_output = _verify_md_output_file(remove_payload, work_dir=work_dir, command_name="number markdown remove")
    removed_text = _read_text_with_long_path(removed_output)
    if removed_text != original:
        raise RuntimeError(f"packaged_number_remove_content_mismatch: {removed_text!r}")
    return added_output, removed_output


def _proofread_line_starts(text: str) -> list[int]:
    starts = [0]
    index = 0
    while index < len(text):
        if text[index] == "\r":
            if index + 1 < len(text) and text[index + 1] == "\n":
                starts.append(index + 2)
                index += 2
                continue
            starts.append(index + 1)
        elif text[index] == "\n":
            starts.append(index + 1)
        index += 1
    return starts


def _verify_proofread_position(
    value: object,
    *,
    expected_offset: int,
    line_starts: list[int],
    label: str,
) -> None:
    if not isinstance(value, dict) or set(value) != {"offset", "line", "column"}:
        raise RuntimeError(f"packaged_proofread_report_{label}_position_invalid:{value}")
    if any(type(value.get(key)) is not int for key in ("offset", "line", "column")):
        raise RuntimeError(f"packaged_proofread_report_{label}_position_invalid:{value}")
    line = bisect_right(line_starts, expected_offset) - 1
    expected = {
        "offset": expected_offset,
        "line": line,
        "column": expected_offset - line_starts[line],
    }
    if value != expected:
        raise RuntimeError(f"packaged_proofread_report_{label}_position_mismatch:expected={expected}:actual={value}")


def _verify_proofread_report(
    report: object,
    *,
    source_path: Path,
    source_bytes: bytes,
    expected_checks: dict[str, bool],
    expect_fixture_issues: bool,
) -> None:
    if not isinstance(report, dict):
        raise RuntimeError(f"packaged_proofread_report_not_object:{report}")
    expected_root_keys = {
        "schema",
        "file",
        "source",
        "location_contract",
        "checks_enabled",
        "issues",
        "summary",
    }
    if set(report) != expected_root_keys:
        raise RuntimeError(f"packaged_proofread_report_root_contract_mismatch:{sorted(report)}")
    if report.get("schema") != "docwen.proofread_report.v2" or report.get("file") != source_path.name:
        raise RuntimeError(f"packaged_proofread_report_identity_mismatch:{report}")
    expected_source = {
        "content_sha256": sha256(source_bytes).hexdigest(),
        "encoding": "utf-8",
        "decode_errors": "replace",
    }
    if report.get("source") != expected_source:
        raise RuntimeError(
            f"packaged_proofread_report_source_mismatch:expected={expected_source}:actual={report.get('source')}"
        )
    if report.get("location_contract") != _PROOFREAD_LOCATION_CONTRACT:
        raise RuntimeError(f"packaged_proofread_report_location_contract_mismatch:{report.get('location_contract')}")
    if report.get("checks_enabled") != expected_checks:
        raise RuntimeError(f"packaged_proofread_report_checks_mismatch:{report.get('checks_enabled')}")

    issues = report.get("issues")
    summary = report.get("summary")
    if not isinstance(issues, list) or not isinstance(summary, dict):
        raise RuntimeError("packaged_proofread_report_issues_or_summary_invalid")

    source_text = source_bytes.decode("utf-8", errors="replace")
    line_starts = _proofread_line_starts(source_text)
    actual_summary: dict[str, int] = {}
    issue_by_signature: dict[tuple[str, str], dict[object, object]] = {}
    for index, issue in enumerate(issues):
        if not isinstance(issue, dict):
            raise RuntimeError(f"packaged_proofread_report_issue_not_object:{index}:{issue}")
        expected_issue_keys = {
            "range",
            "matched_text",
            "error_text",
            "suggestion",
            "error_type",
            "source",
            "rule_key",
        }
        if "fix" in issue:
            expected_issue_keys.add("fix")
        if set(issue) != expected_issue_keys:
            raise RuntimeError(f"packaged_proofread_report_issue_contract_mismatch:{index}:{sorted(issue)}")
        issue_range = issue.get("range")
        if not isinstance(issue_range, dict) or set(issue_range) != {"start", "end"}:
            raise RuntimeError(f"packaged_proofread_report_issue_range_invalid:{index}:{issue_range}")
        start = issue_range.get("start")
        end = issue_range.get("end")
        if not isinstance(start, dict) or not isinstance(end, dict):
            raise RuntimeError(f"packaged_proofread_report_issue_range_invalid:{index}:{issue_range}")
        start_offset = start.get("offset")
        end_offset = end.get("offset")
        if (
            type(start_offset) is not int
            or type(end_offset) is not int
            or start_offset < 0
            or end_offset <= start_offset
            or end_offset > len(source_text)
        ):
            raise RuntimeError(f"packaged_proofread_report_issue_offsets_invalid:{index}:{issue_range}")
        _verify_proofread_position(
            start,
            expected_offset=start_offset,
            line_starts=line_starts,
            label=f"issue_{index}_start",
        )
        _verify_proofread_position(
            end,
            expected_offset=end_offset,
            line_starts=line_starts,
            label=f"issue_{index}_end",
        )
        source_slice = source_text[start_offset:end_offset]
        if issue.get("matched_text") != source_slice or issue.get("error_text") != source_slice:
            raise RuntimeError(
                f"packaged_proofread_report_issue_text_mismatch:{index}:"
                f"slice={source_slice!r}:matched={issue.get('matched_text')!r}:error={issue.get('error_text')!r}"
            )

        issue_source = issue.get("source")
        rule_key = issue.get("rule_key")
        if not isinstance(issue_source, str) or not isinstance(rule_key, str):
            raise RuntimeError(f"packaged_proofread_report_issue_discriminator_invalid:{index}:{issue}")
        actual_summary[rule_key] = actual_summary.get(rule_key, 0) + 1
        issue_by_signature[(issue_source, source_slice)] = issue

        if "fix" in issue:
            fix = issue["fix"]
            if (
                issue_source not in {"typo", "symbol"}
                or not isinstance(fix, dict)
                or set(fix) != {"kind", "replacement", "applicable"}
                or fix.get("kind") != "replace_text"
                or not isinstance(fix.get("replacement"), str)
                or fix.get("applicable") is not True
            ):
                raise RuntimeError(f"packaged_proofread_report_fix_invalid:{index}:{fix}")

    if summary != actual_summary:
        raise RuntimeError(f"packaged_proofread_report_summary_mismatch:expected={actual_summary}:actual={summary}")

    if not expect_fixture_issues:
        if issues or summary:
            raise RuntimeError(f"packaged_proofread_report_expected_successful_empty_result:{report}")
        return

    expected_signatures = {("symbol", "１"), ("symbol", "２"), ("pairing", "（")}
    if set(issue_by_signature) != expected_signatures:
        raise RuntimeError(
            "packaged_proofread_report_fixture_issues_mismatch:"
            f"expected={sorted(expected_signatures)}:actual={sorted(issue_by_signature)}"
        )
    for matched_text, replacement in (("１", "1"), ("２", "2")):
        fix = issue_by_signature[("symbol", matched_text)].get("fix")
        if fix != {"kind": "replace_text", "replacement": replacement, "applicable": True}:
            raise RuntimeError(f"packaged_proofread_report_explicit_replacement_mismatch:{matched_text}:{fix}")
    if "fix" in issue_by_signature[("pairing", "（")]:
        raise RuntimeError("packaged_proofread_report_suggestion_only_issue_exposed_fix")


def _run_proofread_report_validation(
    binary_path: Path,
    *,
    work_dir: Path,
    source_path: Path,
    source_bytes: bytes,
    report_path: Path,
    checks: tuple[str, ...],
    expected_checks: dict[str, bool],
    expect_fixture_issues: bool,
) -> Path:
    check_args = tuple(item for check in checks for item in ("--check", check))
    payload = _load_json_payload(
        _run(
            binary_path,
            "validate",
            str(source_path),
            *check_args,
            "--report",
            str(report_path),
            "--json",
            "--quiet",
            cwd=work_dir,
        ),
        command_name="validate packaged proofread report",
    )
    if payload.get("success") is not True or payload.get("command") != "validate" or payload.get("error") is not None:
        raise RuntimeError(f"packaged_proofread_report_validate_failed:{payload}")
    if _read_bytes_with_long_path(source_path) != source_bytes:
        raise RuntimeError("packaged_proofread_report_source_bytes_changed")
    if not report_path.is_file() or report_path.stat().st_size == 0:
        raise RuntimeError(f"packaged_proofread_report_missing_or_empty:{report_path}")
    data = payload.get("data")
    if not isinstance(data, dict):
        raise RuntimeError(f"packaged_proofread_report_data_missing:{payload}")
    output_path = _resolve_output_file(payload, work_dir=work_dir, command_name="validate packaged proofread report")
    if output_path != report_path.resolve():
        raise RuntimeError(
            f"packaged_proofread_report_output_mismatch:expected={report_path.resolve()}:actual={output_path}"
        )
    try:
        report = json.loads(_read_text_with_long_path(report_path))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"packaged_proofread_report_invalid_json:{report_path}") from exc
    details = data.get("details")
    if not isinstance(details, dict) or details.get("proofread") != report:
        raise RuntimeError(f"packaged_proofread_report_inline_projection_mismatch:{details}")
    _verify_proofread_report(
        report,
        source_path=source_path,
        source_bytes=source_bytes,
        expected_checks=expected_checks,
        expect_fixture_issues=expect_fixture_issues,
    )
    return report_path


def _run_optional_proofread_report_smoke(binary_path: Path, *, work_dir: Path) -> tuple[Path, Path]:
    source_path = work_dir / "proofread-report-2.0-coordinates.md"
    source_bytes = _write_proofread_report_fixture(source_path)
    populated_report = _run_proofread_report_validation(
        binary_path,
        work_dir=work_dir,
        source_path=source_path,
        source_bytes=source_bytes,
        report_path=work_dir / "proofread-report-2.0-issues.json",
        checks=("symbol", "punct"),
        expected_checks={
            "symbol_pairing": True,
            "symbol_correction": True,
            "typos_rule": False,
            "sensitive_word": False,
        },
        expect_fixture_issues=True,
    )
    empty_report = _run_proofread_report_validation(
        binary_path,
        work_dir=work_dir,
        source_path=source_path,
        source_bytes=source_bytes,
        report_path=work_dir / "proofread-report-2.0-empty.json",
        checks=("none",),
        expected_checks={
            "symbol_pairing": False,
            "symbol_correction": False,
            "typos_rule": False,
            "sensitive_word": False,
        },
        expect_fixture_issues=False,
    )
    return populated_report, empty_report


def _run_optional_successful_warning_smoke(
    binary_path: Path,
    *,
    work_dir: Path,
    input_path: Path,
    action: str,
    expected_code: str,
    expected_message: str = "",
) -> Path:
    """Verify successful warning projection in packaged JSON and text modes."""
    json_output_dir = work_dir / "warning_json"
    text_output_dir = work_dir / "warning_text"
    json_output_dir.mkdir()
    text_output_dir.mkdir()

    payload = _load_json_payload(
        _run(
            binary_path,
            "convert",
            str(input_path),
            "--to",
            "md",
            "--optimization",
            action,
            "--output",
            str(json_output_dir),
            "--json",
            cwd=work_dir,
        ),
        command_name="successful warning JSON convert",
    )
    if payload.get("success") is not True:
        raise RuntimeError(f"successful warning JSON convert returned success=false: {payload}")
    warnings = payload.get("warnings")
    if not isinstance(warnings, list):
        raise RuntimeError(f"packaged_successful_warning_list_missing: {payload}")
    matching = [
        item
        for item in warnings
        if isinstance(item, dict) and item.get("level") == "warning" and item.get("code") == expected_code
    ]
    if len(matching) != 1:
        raise RuntimeError(
            f"packaged_successful_warning_code_mismatch: expected one {expected_code!r}, got {warnings!r}"
        )
    message = matching[0].get("message")
    if not isinstance(message, str) or not message:
        raise RuntimeError(f"packaged_successful_warning_message_missing: {matching[0]!r}")
    if expected_message and message != expected_message:
        raise RuntimeError(
            f"packaged_successful_warning_message_mismatch: expected {expected_message!r}, got {message!r}"
        )
    verified_json_output = _verify_md_output_file(
        payload,
        work_dir=work_dir,
        command_name="successful warning JSON convert",
    )

    text_proc = _run(
        binary_path,
        "convert",
        str(input_path),
        "--to",
        "md",
        "--optimization",
        action,
        "--output",
        str(text_output_dir),
        cwd=work_dir,
    )
    _verify_process_succeeded(text_proc, command_name="successful warning text convert")
    if expected_code in text_proc.stdout:
        raise RuntimeError(f"packaged_successful_warning_leaked_to_stdout: {text_proc.stdout}")
    if expected_code not in text_proc.stderr:
        raise RuntimeError(f"packaged_successful_warning_missing_from_stderr: {text_proc.stderr}")
    if expected_message and expected_message not in text_proc.stderr:
        raise RuntimeError(f"packaged_successful_warning_message_missing_from_stderr: {text_proc.stderr}")

    expected_bytes = _read_bytes_with_long_path(verified_json_output)
    text_outputs = _find_markdown_outputs_with_long_path(text_output_dir)
    matching_outputs = [path for path in text_outputs if _read_bytes_with_long_path(path) == expected_bytes]
    if len(matching_outputs) != 1:
        raise RuntimeError(
            "packaged_successful_warning_text_output_mismatch: "
            f"expected_one_byte_identical_primary; candidates={[str(path) for path in text_outputs]}"
        )
    return verified_json_output


def _normalized_field_instruction(parts: list[str]) -> str:
    return " ".join("".join(parts).split())


def _word_complex_fields(root: ElementTree.Element) -> list[tuple[str, str]]:
    fields: list[tuple[str, str]] = []
    for paragraph in root.iter(f"{_WORD_TAG}p"):
        instruction_parts: list[str] | None = None
        result_parts: list[str] = []
        in_result = False
        for element in paragraph.iter():
            if element.tag == f"{_WORD_TAG}fldChar":
                field_type = element.get(f"{_WORD_TAG}fldCharType")
                if field_type == "begin":
                    instruction_parts = []
                    result_parts = []
                    in_result = False
                elif instruction_parts is not None and field_type == "separate":
                    in_result = True
                elif instruction_parts is not None and field_type == "end":
                    fields.append((_normalized_field_instruction(instruction_parts), "".join(result_parts)))
                    instruction_parts = None
                    result_parts = []
                    in_result = False
            elif instruction_parts is not None and element.tag == f"{_WORD_TAG}instrText":
                instruction_parts.append(element.text or "")
            elif instruction_parts is not None and in_result and element.tag == f"{_WORD_TAG}t":
                result_parts.append(element.text or "")
    return fields


def verify_machine_document_semantics_docx(path: Path) -> None:
    """Fail closed unless a Machine-produced DOCX carries the v4 exact-two probe."""

    try:
        with _zipfile_with_long_path(path) as archive:
            document_xml = archive.read("word/document.xml")
            media_names = tuple(name for name in archive.namelist() if name.startswith("word/media/"))
    except (OSError, KeyError, zipfile.BadZipFile) as exc:
        raise RuntimeError(f"packaged_machine_semantics_docx_invalid:{path}") from exc
    try:
        root = ElementTree.fromstring(document_xml)
    except ElementTree.ParseError as exc:
        raise RuntimeError("packaged_machine_semantics_document_xml_invalid") from exc

    fields = _word_complex_fields(root)
    for counter in ("Figure", "Table", "Equation", "Code"):
        sequence_fields = [
            (instruction, result) for instruction, result in fields if instruction.startswith(f"SEQ {counter} ")
        ]
        if len(sequence_fields) != 1:
            raise RuntimeError(f"packaged_machine_exact_two_seq_invalid:{counter}:{fields}")
    reference_fields = [(instruction, result) for instruction, result in fields if instruction.startswith("REF DW_T_")]
    if len(reference_fields) != 2:
        raise RuntimeError(f"packaged_machine_exact_two_reference_count_invalid:{reference_fields}")
    if len({result for _instruction, result in reference_fields}) != 1:
        raise RuntimeError("packaged_machine_exact_two_reference_cached_mismatch")
    bookmark_names = [element.get(f"{_WORD_TAG}name", "") for element in root.iter(f"{_WORD_TAG}bookmarkStart")]
    target_bookmarks = [name for name in bookmark_names if name.startswith("DW_T_")]
    if len(target_bookmarks) != 5 or len(set(target_bookmarks)) != 5:
        raise RuntimeError(f"packaged_machine_exact_two_target_bookmarks_invalid:{target_bookmarks}")
    for reference_target in (reference_fields[0][0].split()[1], reference_fields[1][0].split()[1]):
        if reference_target not in bookmark_names:
            raise RuntimeError(f"packaged_machine_exact_two_reference_bookmark_missing:{reference_target}")

    caption_sdt = [
        element
        for element in root.iter()
        if element.tag == f"{_WORD_TAG}tag"
        and (element.get(f"{_WORD_TAG}val") or "").startswith("docwen-numbering-occurrence-v1:")
    ]
    if caption_sdt:
        raise RuntimeError("packaged_machine_exact_two_unexpected_disabled_occurrence")
    citation_fields = [(instruction, result) for instruction, result in fields if instruction.startswith("CITATION ")]
    if len(citation_fields) != 1:
        raise RuntimeError(f"packaged_machine_exact_two_citation_field_invalid:{citation_fields}")
    if "One (2026)" not in "".join(element.text or "" for element in root.iter(f"{_WORD_TAG}t")):
        raise RuntimeError("packaged_machine_exact_two_citation_cache_missing")
    visible_text = "".join(element.text or "" for element in root.iter(f"{_WORD_TAG}t"))
    for token in ("Architecture", "System overview", "Results", "Entry point"):
        if token not in visible_text:
            raise RuntimeError(f"packaged_machine_exact_two_visible_text_missing:{token}")
    for counter in ("Figure", "Table", "Equation", "Code"):
        caption_pattern = re.compile(rf"{counter}\s*1")
        if caption_pattern.search(visible_text) is None:
            raise RuntimeError(f"packaged_machine_exact_two_caption_text_missing:{counter}:{visible_text[:400]}")
    if not list(root.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath")):
        raise RuntimeError("packaged_machine_exact_two_equation_omml_missing")
    if not media_names:
        raise RuntimeError("packaged_machine_exact_two_media_missing")
        raise RuntimeError("packaged_machine_semantics_image_missing")


def verify_machine_document_semantics_markdown(markdown: str) -> None:
    """Verify the canonical Markdown projection of the packaged exact-two probe."""

    required_tokens = (
        "# Architecture ^h-7f3a",
        "Figure: System overview ^system-overview",
        "Table: Results ^results-main",
        "Equation: ^energy-main",
        "Code: Entry point ^entry-main",
        "Stable: @[[#^h-7f3a]] and @[[#^system-overview|System overview]].",
        "Ordinary: [[#^system-overview]] and ![[Guide#^h-7f3a]].",
        "Citation: @cite-one.",
    )
    missing = [token for token in required_tokens if token not in markdown]
    if missing:
        raise RuntimeError(f"packaged_machine_semantics_markdown_missing:{missing}")
    if "system.png" not in markdown:
        raise RuntimeError("packaged_machine_semantics_markdown_image_missing")


MACHINE_EXACT_TWO_NOTE_ROUNDTRIP_TOKENS = (
    "Notes: default[^1], explicit[^2], first endnote[^endnote:1], second endnote[^endnote:2].",
    "[^1]: Default footnote.",
    "[^2]: Explicit footnote.",
    "[^endnote:1]: Canonical endnote.",
    "[^endnote:2]: Second endnote.",
)


def verify_machine_note_domains_markdown(markdown: str) -> None:
    """Verify note-domain preservation in the packaged exact-two round trip."""

    missing = [token for token in MACHINE_EXACT_TWO_NOTE_ROUNDTRIP_TOKENS if token not in markdown]
    if missing:
        raise RuntimeError(f"packaged_machine_note_domains_markdown_missing:{missing}")


def _machine_frame(payload: dict[str, Any]) -> bytes:
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body


def _read_machine_frame(stream: IO[bytes]) -> dict[str, Any]:
    header = bytearray()
    while not header.endswith(b"\r\n\r\n"):
        byte = stream.read(1)
        if not byte:
            raise RuntimeError("packaged_machine_protocol_unexpected_eof")
        header.extend(byte)
        if len(header) > 64:
            raise RuntimeError("packaged_machine_protocol_invalid_header")
    match = re.fullmatch(rb"Content-Length: ([1-9][0-9]*)\r\n\r\n", bytes(header))
    if match is None:
        raise RuntimeError(f"packaged_machine_protocol_invalid_header:{bytes(header)!r}")
    content_length = int(match.group(1))
    body = stream.read(content_length)
    if len(body) != content_length:
        raise RuntimeError("packaged_machine_protocol_length_mismatch")
    try:
        payload = json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RuntimeError("packaged_machine_protocol_invalid_json") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("packaged_machine_protocol_non_object")
    return payload


class _MachineSmokeResources:
    """Own process and temp resources across every Machine smoke exit path."""

    def __init__(self) -> None:
        self.process: subprocess.Popen[bytes] | None = None
        self.physical_temp: tempfile.TemporaryDirectory[str] | None = None

    def cleanup(self) -> list[str]:
        errors: list[str] = []
        process = self.process
        if process is not None:
            if process.stdin is not None and not process.stdin.closed:
                with contextlib.suppress(OSError):
                    process.stdin.close()
            if process.poll() is None:
                try:
                    process.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    with contextlib.suppress(OSError):
                        process.terminate()
                    try:
                        process.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        with contextlib.suppress(OSError):
                            process.kill()
                        try:
                            process.wait(timeout=5)
                        except (OSError, subprocess.TimeoutExpired) as exc:
                            errors.append(f"process_wait:{type(exc).__name__}:{exc}")
                except OSError as exc:
                    errors.append(f"process_wait:{type(exc).__name__}:{exc}")
            for stream in (process.stdout, process.stderr):
                if stream is not None and not stream.closed:
                    with contextlib.suppress(OSError):
                        stream.close()
            self.process = None
        if self.physical_temp is not None:
            try:
                self.physical_temp.cleanup()
            except OSError as exc:
                errors.append(f"physical_temp:{type(exc).__name__}:{exc}")
            self.physical_temp = None
        return errors


def _run_machine_protocol_smoke(binary_path: Path | Sequence[str], *, work_dir: Path) -> Path:
    resources = _MachineSmokeResources()
    failure: BaseException | None = None
    try:
        return _run_machine_protocol_smoke_impl(binary_path, work_dir=work_dir, resources=resources)
    except BaseException as exc:
        failure = exc
        raise
    finally:
        cleanup_errors = resources.cleanup()
        if cleanup_errors:
            message = f"packaged_machine_protocol_cleanup_failed:{cleanup_errors}"
            if failure is not None:
                failure.add_note(message)
            else:
                raise RuntimeError(message)


def _run_machine_protocol_smoke_impl(
    binary_path: Path | Sequence[str],
    *,
    work_dir: Path,
    resources: _MachineSmokeResources,
) -> Path:
    """Execute packaged single- and multi-artifact framed-protocol tasks."""

    source = work_dir / "machine-source" / "机器协议 输入.md"
    source.parent.mkdir()
    source.write_text(MACHINE_DOCUMENT_SEMANTICS_FIXTURE, encoding="utf-8")
    decoy_image = source.parent / "assets" / "机器协议-语义.png"
    decoy_image.parent.mkdir()
    decoy_image.write_bytes(b"undeclared physical sibling decoy: must not be read")
    source_bytes = _read_bytes_with_long_path(source)
    neutral_document = work_dir / "machine-exact-two" / "neutral-document.json"
    neutral_document.parent.mkdir()
    neutral_document.write_text(
        json.dumps(MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    numbering_plan = work_dir / "machine-exact-two" / "numbering-export-plan.json"
    numbering_plan.write_text(
        json.dumps(MACHINE_EXACT_TWO_NUMBERING_PLAN, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    neutral_bytes = _read_bytes_with_long_path(neutral_document)
    numbering_plan_bytes = _read_bytes_with_long_path(numbering_plan)
    staging = work_dir / "machine-staging"
    staging.mkdir()
    semantic_reverse_staging = work_dir / "machine-semantic-reverse-staging"
    semantic_reverse_staging.mkdir()
    ocr_source = work_dir / "机器协议 OCR.png"
    _write_ocr_png(ocr_source)
    ocr_source_bytes = _read_bytes_with_long_path(ocr_source)
    ocr_staging = work_dir / "machine-ocr-staging"
    ocr_staging.mkdir()
    validation_staging = work_dir / "machine-validation-staging"
    validation_staging.mkdir()
    merge_staging = work_dir / "machine-merge-staging"
    merge_staging.mkdir()
    merge_pdf_a = work_dir / "machine-merge-a.pdf"
    merge_pdf_b = work_dir / "machine-merge-b.pdf"
    _write_pymupdf_layout_pdf(merge_pdf_a)
    _write_pymupdf_layout_pdf(merge_pdf_b)
    merge_pdf_a_bytes = _read_bytes_with_long_path(merge_pdf_a)
    merge_pdf_b_bytes = _read_bytes_with_long_path(merge_pdf_b)
    physical_pdf = work_dir / "machine-physical-pages.pdf"
    _write_physical_page_pdf(physical_pdf)
    physical_tiff = work_dir / "machine-physical-frames.tiff"
    _write_physical_page_tiff(physical_tiff)
    physical_ofd = work_dir / "machine-physical-pages.ofd"
    _write_physical_page_ofd(physical_ofd)
    physical_xps = work_dir / "machine-physical-pages.xps"
    _write_physical_page_xps(physical_xps)
    # Runtime workspaces contain producer-owned nested staging paths. Keep the
    # runtime temp root short even though the surrounding verifier workspace is
    # deliberately 200+ characters for separate long-path coverage.
    resources.physical_temp = tempfile.TemporaryDirectory(prefix="dw-physical-")
    physical_governed_root = Path(resources.physical_temp.name)
    (physical_governed_root / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    physical_runtime_temp = physical_governed_root / "temp" / "runtime"
    physical_system_temp = physical_governed_root / "temp" / "system"
    physical_runtime_temp.mkdir(parents=True)
    physical_system_temp.mkdir()
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["DOCWEN_CONFIG_DIR"] = str(work_dir / "config_home")
    env["DOCWEN_LOG_DIR"] = str(work_dir / "log_home")
    env["DOCWEN_LOG_TO_TEMP"] = ""
    env["DOCWEN_WORKSPACE_ROOT"] = str(physical_governed_root)
    env["TEMP"] = str(physical_system_temp)
    env["TMP"] = str(physical_system_temp)
    env["TMPDIR"] = str(physical_system_temp)
    machine_command = [str(binary_path)] if isinstance(binary_path, (str, Path)) else list(binary_path)
    process = subprocess.Popen(
        [*machine_command, "serve", "--stdio"],
        cwd=work_dir,
        env=env,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    resources.process = process
    if process.stdin is None or process.stdout is None or process.stderr is None:
        raise RuntimeError("packaged_machine_protocol_stdio_unavailable")
    machine_stdin = process.stdin
    machine_stdout = process.stdout
    machine_stderr = process.stderr

    def exchange(request: dict[str, Any]) -> dict[str, Any]:
        machine_stdin.write(_machine_frame(request))
        machine_stdin.flush()
        return _read_machine_frame(machine_stdout)

    initialize = exchange(
        {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
                "client": {"name": "packaged-verifier", "version": "1.0.0"},
                "features": {"progress": True, "cancellation": True},
            },
        }
    )
    result = initialize.get("result")
    if not isinstance(result, dict) or result.get("artifact_bundle_schema") != "docwen.artifact_bundle.v2":
        raise RuntimeError(f"packaged_machine_protocol_initialize_invalid:{initialize}")

    discovery = exchange({"jsonrpc": "2.0", "id": 2, "method": "capability/list", "params": {}})
    discovery_result = discovery.get("result")
    capability_items = discovery_result.get("capabilities") if isinstance(discovery_result, dict) else None
    if not isinstance(capability_items, list):
        raise RuntimeError(f"packaged_machine_protocol_discovery_invalid:{discovery}")
    capability_ids = {
        capability_id
        for item in capability_items
        if isinstance(item, dict)
        for capability_id in (item.get("capability_id"),)
        if isinstance(capability_id, str)
    }
    expected_capability_ids = {
        "convert.markdown.to_docx",
        "convert.markdown.to_xlsx",
        "convert.docx.to_markdown",
        "convert.pdf.to_markdown",
        "convert.ofd.to_markdown",
        "convert.xps.to_markdown",
        "convert.tiff.to_markdown",
        "convert.xlsx.to_markdown",
        "convert.xlsx.to_csv",
        "render.pdf.to_png",
        "split.pdf.every_page",
        "convert.png.to_ocr_markdown",
        "convert.markdown_tables.to_csv",
        "convert.tiff_frames.to_png",
        "validate.markdown",
        "transform.markdown.heading_numbering",
        "merge.pdf.documents",
        "split.pdf.partition",
        "merge.xlsx.tables",
        "merge.images.to_tiff",
    }
    if capability_ids != expected_capability_ids:
        raise RuntimeError(f"packaged_machine_protocol_capabilities_mismatch:{sorted(capability_ids)}")
    if len(capability_items) != len(expected_capability_ids):
        raise RuntimeError("packaged_machine_protocol_capability_count_mismatch")
    capability_by_id = {
        item["capability_id"]: item
        for item in capability_items
        if isinstance(item, dict) and isinstance(item.get("capability_id"), str)
    }
    expected_semantic_limitations = list(MACHINE_DOCUMENT_SEMANTICS_LIMITATIONS)
    expected_resolved_document_limitations = list(MACHINE_RESOLVED_DOCUMENT_LIMITATIONS)
    for capability_id, expected in (
        ("convert.markdown.to_docx", expected_resolved_document_limitations),
        ("convert.docx.to_markdown", expected_semantic_limitations),
    ):
        actual_limitations = capability_by_id[capability_id].get("limitations")
        if actual_limitations != expected:
            raise RuntimeError(
                f"packaged_machine_protocol_semantic_limitations_mismatch:{capability_id}:{actual_limitations}"
            )
    markdown_capability = capability_by_id.get("convert.markdown.to_docx")
    expected_markdown_input_shape = {
        "slots": [
            {
                "role": "neutral_document",
                "kind": "document",
                "media_types": [MACHINE_RESOLVED_DOCUMENT_MEDIA_TYPE],
                "min_items": 1,
                "max_items": 1,
            },
            {
                "role": "numbering_export_plan",
                "kind": "resource",
                "media_types": [MACHINE_NUMBERING_EXPORT_PLAN_MEDIA_TYPE],
                "min_items": 1,
                "max_items": 1,
            },
        ],
        "undeclared_roles": "reject",
    }
    if (
        not isinstance(markdown_capability, dict)
        or markdown_capability.get("input_shape") != expected_markdown_input_shape
    ):
        raise RuntimeError(
            "packaged_machine_protocol_markdown_input_shape_mismatch:"
            f"{markdown_capability.get('input_shape') if isinstance(markdown_capability, dict) else markdown_capability}"
        )
    physical_capability_ids = (
        "convert.pdf.to_markdown",
        "convert.ofd.to_markdown",
        "convert.xps.to_markdown",
        "convert.tiff.to_markdown",
    )
    expected_physical_shape = {
        "cardinality": "many",
        "artifact_kinds": ["document", "fragment", "resource"],
        "relation_types": ["fragment_of", "resource_of"],
        "atomic_bundle": True,
        "relation_payloads": ["page_fragment", "page_resource"],
    }
    fixed_layout_properties = {
        "recognize_text": {"type": "boolean", "default": False},
        "preserve_resources": {"type": "boolean", "default": True},
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
        },
        "image_mode": {"type": "string", "enum": ["file"], "default": "file"},
        "render_dpi": {"type": "integer", "minimum": 72, "maximum": 600, "default": 200},
    }
    fixed_layout_options = {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "properties": fixed_layout_properties,
        "required": [],
        "additionalProperties": False,
    }
    tiff_options = {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "properties": {
            "recognize_text": {"type": "boolean", "default": False},
            "preserve_resources": {"type": "boolean", "default": True},
            "ocr_language": fixed_layout_properties["ocr_language"],
        },
        "required": [],
        "additionalProperties": False,
    }
    expected_physical_capabilities = {
        "convert.pdf.to_markdown": (
            "application/pdf",
            fixed_layout_options,
            [
                {"dependency_id": "python.pymupdf4llm", "required": True, "available": True},
                {"dependency_id": "python.rapidocr", "required": False, "available": True},
            ],
            [
                *MACHINE_PHYSICAL_PAGE_LIMITATIONS,
                {
                    "severity": "warning",
                    "code": "runtime_route_limitation",
                    "message": "OCR options require the optional RapidOCR capability",
                },
            ],
        ),
        "convert.ofd.to_markdown": (
            "application/vnd.ofd",
            fixed_layout_options,
            [
                {"dependency_id": "python.easyofd", "required": True, "available": True},
                {"dependency_id": "python.pymupdf4llm", "required": True, "available": True},
                {"dependency_id": "python.rapidocr", "required": False, "available": True},
            ],
            [
                *MACHINE_PHYSICAL_PAGE_LIMITATIONS,
                {
                    "severity": "warning",
                    "code": "runtime_route_limitation",
                    "message": "OCR options require the optional RapidOCR capability",
                },
                {
                    "severity": "warning",
                    "code": "runtime_route_limitation",
                    "message": "OFD input is normalized to PDF before the selected route runs",
                },
            ],
        ),
        "convert.xps.to_markdown": (
            "application/vnd.ms-xpsdocument",
            fixed_layout_options,
            [
                {"dependency_id": "python.fitz", "required": True, "available": True},
                {"dependency_id": "python.pymupdf4llm", "required": True, "available": True},
                {"dependency_id": "python.rapidocr", "required": False, "available": True},
            ],
            [
                *MACHINE_PHYSICAL_PAGE_LIMITATIONS,
                {
                    "severity": "warning",
                    "code": "runtime_route_limitation",
                    "message": "OCR options require the optional RapidOCR capability",
                },
                {
                    "severity": "warning",
                    "code": "runtime_route_limitation",
                    "message": "XPS input is normalized to PDF before the selected route runs",
                },
            ],
        ),
        "convert.tiff.to_markdown": (
            "image/tiff",
            tiff_options,
            [
                {"dependency_id": "python.pillow", "required": True, "available": True},
                {"dependency_id": "python.rapidocr", "required": False, "available": True},
            ],
            list(MACHINE_PHYSICAL_PAGE_LIMITATIONS),
        ),
    }
    for capability_id in physical_capability_ids:
        capability = capability_by_id.get(capability_id)
        media_type, options_schema, dependencies, limitations = expected_physical_capabilities[capability_id]
        expected_input_shape = {
            "slots": [
                {
                    "role": "source",
                    "kind": "resource",
                    "media_types": [media_type],
                    "min_items": 1,
                    "max_items": 1,
                }
            ],
            "undeclared_roles": "reject",
        }
        expected_capability = {
            "capability_id": capability_id,
            "operation": "convert",
            "input_shape": expected_input_shape,
            "output_media_types": ["text/markdown"],
            "output_shape": expected_physical_shape,
            "options_schema": options_schema,
            "availability": "available",
            "dependencies": dependencies,
            "limitations": limitations,
        }
        if capability != expected_capability:
            raise RuntimeError(f"packaged_physical_page_capability_shape_invalid:{capability_id}:{capability}")
        serialized_capability = json.dumps(capability, ensure_ascii=False).casefold()
        if any(token in serialized_capability for token in ("page_nodes", "pkwf", "wenleaf")):
            raise RuntimeError(f"packaged_physical_page_capability_consumer_policy_leak:{capability_id}")

    planned = exchange(
        {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "task/plan",
            "params": {
                "capability_id": "convert.markdown.to_docx",
                "inputs": [
                    {
                        "input_id": "input.neutral-document",
                        "kind": "document",
                        "role": "neutral_document",
                        "logical_path": "inputs/document.resolved.json",
                        "locator": {"kind": "local_path", "path": str(neutral_document)},
                        "media_type": MACHINE_RESOLVED_DOCUMENT_MEDIA_TYPE,
                        "size_bytes": len(neutral_bytes),
                        "sha256": hashlib.sha256(neutral_bytes).hexdigest(),
                    },
                    {
                        "input_id": "input.numbering-export-plan",
                        "kind": "resource",
                        "role": "numbering_export_plan",
                        "logical_path": "inputs/numbering-export-plan.json",
                        "locator": {"kind": "local_path", "path": str(numbering_plan)},
                        "media_type": MACHINE_NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
                        "size_bytes": len(numbering_plan_bytes),
                        "sha256": hashlib.sha256(numbering_plan_bytes).hexdigest(),
                    },
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        }
    )
    plan_result = planned.get("result")
    if not isinstance(plan_result, dict):
        raise RuntimeError(f"packaged_machine_protocol_plan_invalid:{planned}")
    plan_id = plan_result.get("plan_id")
    if not isinstance(plan_id, str):
        raise RuntimeError(f"packaged_machine_protocol_plan_invalid:{planned}")
    if plan_result.get("limitations") != expected_resolved_document_limitations:
        raise RuntimeError(f"packaged_machine_protocol_plan_limitations_mismatch:{planned}")
    accepted = exchange(
        {
            "jsonrpc": "2.0",
            "id": 4,
            "method": "task/execute",
            "params": {"plan_id": plan_id},
        }
    )
    accepted_result = accepted.get("result")
    task_id = accepted_result.get("task_id") if isinstance(accepted_result, dict) else None
    accepted_state = accepted_result.get("state") if isinstance(accepted_result, dict) else None
    if not isinstance(task_id, str) or accepted_state != "accepted":
        raise RuntimeError(f"packaged_machine_protocol_accept_invalid:{accepted}")

    terminal: dict[str, Any] | None = None
    while terminal is None:
        notification = _read_machine_frame(machine_stdout)
        if notification.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            terminal = notification

    ocr_planned = exchange(
        {
            "jsonrpc": "2.0",
            "id": 5,
            "method": "task/plan",
            "params": {
                "capability_id": "convert.png.to_ocr_markdown",
                "inputs": [
                    {
                        "input_id": "input.ocr",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": "images/机器协议 OCR.png",
                        "locator": {"kind": "local_path", "path": str(ocr_source)},
                        "media_type": "image/png",
                        "size_bytes": len(ocr_source_bytes),
                        "sha256": hashlib.sha256(ocr_source_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(ocr_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        }
    )
    ocr_plan_result = ocr_planned.get("result")
    ocr_plan_id = ocr_plan_result.get("plan_id") if isinstance(ocr_plan_result, dict) else None
    if not isinstance(ocr_plan_id, str):
        raise RuntimeError(f"packaged_machine_protocol_ocr_plan_invalid:{ocr_planned}")
    ocr_accepted = exchange(
        {
            "jsonrpc": "2.0",
            "id": 6,
            "method": "task/execute",
            "params": {"plan_id": ocr_plan_id},
        }
    )
    ocr_accepted_result = ocr_accepted.get("result")
    ocr_task_id = ocr_accepted_result.get("task_id") if isinstance(ocr_accepted_result, dict) else None
    ocr_accepted_state = ocr_accepted_result.get("state") if isinstance(ocr_accepted_result, dict) else None
    if not isinstance(ocr_task_id, str) or ocr_accepted_state != "accepted":
        raise RuntimeError(f"packaged_machine_protocol_ocr_accept_invalid:{ocr_accepted}")
    ocr_terminal: dict[str, Any] | None = None
    while ocr_terminal is None:
        notification = _read_machine_frame(machine_stdout)
        if notification.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            ocr_terminal = notification

    def execute_additional_task(
        *,
        plan_request_id: int,
        execute_request_id: int,
        capability_id: str,
        inputs: list[dict[str, Any]],
        output_root: Path,
        options: dict[str, Any],
    ) -> tuple[str, dict[str, Any]]:
        task_plan = exchange(
            {
                "jsonrpc": "2.0",
                "id": plan_request_id,
                "method": "task/plan",
                "params": {
                    "capability_id": capability_id,
                    "inputs": inputs,
                    "output": {
                        "staging_root": {"kind": "local_path", "path": str(output_root)},
                        "staging_policy": "require_empty",
                    },
                    "options": options,
                },
            }
        )
        task_plan_result = task_plan.get("result")
        if not isinstance(task_plan_result, dict):
            raise RuntimeError(f"packaged_machine_protocol_additional_plan_invalid:{task_plan}")
        task_plan_id = task_plan_result.get("plan_id")
        if not isinstance(task_plan_id, str):
            raise RuntimeError(f"packaged_machine_protocol_additional_plan_invalid:{task_plan}")
        if (
            capability_id in {"convert.markdown.to_docx", "convert.docx.to_markdown"}
            and task_plan_result.get("limitations") != expected_semantic_limitations
        ):
            raise RuntimeError(f"packaged_machine_protocol_additional_plan_limitations_mismatch:{task_plan}")
        if capability_id in physical_capability_ids and task_plan_result.get("limitations") != list(
            MACHINE_PHYSICAL_PAGE_LIMITATIONS
        ):
            raise RuntimeError(f"packaged_physical_page_plan_limitations_mismatch:{task_plan}")
        task_acceptance = exchange(
            {
                "jsonrpc": "2.0",
                "id": execute_request_id,
                "method": "task/execute",
                "params": {"plan_id": task_plan_id},
            }
        )
        task_acceptance_result = task_acceptance.get("result")
        additional_task_id = task_acceptance_result.get("task_id") if isinstance(task_acceptance_result, dict) else None
        if (
            not isinstance(additional_task_id, str)
            or not isinstance(task_acceptance_result, dict)
            or task_acceptance_result.get("state") != "accepted"
        ):
            raise RuntimeError(f"packaged_machine_protocol_additional_accept_invalid:{task_acceptance}")
        additional_terminal: dict[str, Any] | None = None
        while additional_terminal is None:
            notification = _read_machine_frame(machine_stdout)
            if notification.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
                additional_terminal = notification
        return additional_task_id, additional_terminal

    if terminal.get("method") != "task/completed":
        raise RuntimeError(f"packaged_machine_protocol_terminal_invalid:{terminal}")
    semantic_params = terminal.get("params")
    semantic_bundle = semantic_params.get("bundle") if isinstance(semantic_params, dict) else None
    semantic_artifacts = semantic_bundle.get("artifacts") if isinstance(semantic_bundle, dict) else None
    semantic_documents = (
        [item for item in semantic_artifacts if isinstance(item, dict) and item.get("kind") == "document"]
        if isinstance(semantic_artifacts, list)
        else []
    )
    if (
        not isinstance(semantic_bundle, dict)
        or semantic_bundle.get("task_id") != task_id
        or not isinstance(semantic_artifacts, list)
        or len(semantic_artifacts) != 1
        or len(semantic_documents) != 1
        or semantic_bundle.get("relations") != []
    ):
        raise RuntimeError(f"packaged_machine_protocol_semantic_artifact_invalid:{semantic_bundle}")
    expected_entries = [
        {
            "artifact_id": semantic_documents[0].get("artifact_id"),
            "role": "primary",
            "ordinal": 0,
            "preferred": True,
        }
    ]
    if semantic_bundle.get("entries") != expected_entries:
        raise RuntimeError(f"packaged_machine_protocol_semantic_entries_invalid:{semantic_bundle}")
    semantic_locator = semantic_documents[0].get("locator")
    if not isinstance(semantic_locator, str) or "\\" in semantic_locator or ".." in semantic_locator.split("/"):
        raise RuntimeError(f"packaged_machine_protocol_semantic_locator_invalid:{semantic_locator}")
    semantic_output = staging / Path(semantic_locator)
    semantic_output_bytes = _read_bytes_with_long_path(semantic_output)
    if len(semantic_output_bytes) != semantic_documents[0].get("size_bytes") or hashlib.sha256(
        semantic_output_bytes
    ).hexdigest() != semantic_documents[0].get("sha256"):
        raise RuntimeError("packaged_machine_protocol_semantic_integrity_mismatch")
    verify_machine_document_semantics_docx(semantic_output)
    semantic_reverse_task_id, semantic_reverse_terminal = execute_additional_task(
        plan_request_id=7,
        execute_request_id=8,
        capability_id="convert.docx.to_markdown",
        inputs=[
            {
                "input_id": "input.semantic-docx",
                "kind": "document",
                "role": "source",
                "logical_path": "documents/semantic-output.docx",
                "locator": {"kind": "local_path", "path": str(semantic_output)},
                "media_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "size_bytes": len(semantic_output_bytes),
                "sha256": hashlib.sha256(semantic_output_bytes).hexdigest(),
            }
        ],
        output_root=semantic_reverse_staging,
        options={},
    )

    validation_task_id, validation_terminal = execute_additional_task(
        plan_request_id=9,
        execute_request_id=10,
        capability_id="validate.markdown",
        inputs=[
            {
                "input_id": "input.validation",
                "kind": "document",
                "role": "source",
                "logical_path": "documents/机器协议 输入.md",
                "locator": {"kind": "local_path", "path": str(source)},
                "media_type": "text/markdown",
                "size_bytes": len(source_bytes),
                "sha256": hashlib.sha256(source_bytes).hexdigest(),
            }
        ],
        output_root=validation_staging,
        options={},
    )
    merge_task_id, merge_terminal = execute_additional_task(
        plan_request_id=11,
        execute_request_id=12,
        capability_id="merge.pdf.documents",
        inputs=[
            {
                "input_id": "input.merge.1",
                "kind": "resource",
                "role": "source",
                "logical_path": "documents/merge-a.pdf",
                "locator": {"kind": "local_path", "path": str(merge_pdf_a)},
                "media_type": "application/pdf",
                "size_bytes": len(merge_pdf_a_bytes),
                "sha256": hashlib.sha256(merge_pdf_a_bytes).hexdigest(),
            },
            {
                "input_id": "input.merge.2",
                "kind": "resource",
                "role": "source",
                "logical_path": "documents/merge-b.pdf",
                "locator": {"kind": "local_path", "path": str(merge_pdf_b)},
                "media_type": "application/pdf",
                "size_bytes": len(merge_pdf_b_bytes),
                "sha256": hashlib.sha256(merge_pdf_b_bytes).hexdigest(),
            },
        ],
        output_root=merge_staging,
        options={},
    )

    def execute_physical_task(
        *,
        ordinal: int,
        capability_id: str,
        source_path: Path,
        media_type: str,
        ocr_enabled: bool,
        keep_images: bool,
        inject_unresolved_resource: bool = False,
    ) -> tuple[str, dict[str, Any], Path]:
        if list(physical_runtime_temp.glob("dw-*")):
            raise RuntimeError("packaged_physical_page_workspace_residue_before_task")
        output_root = work_dir / f"machine-physical-{ordinal:02d}-staging"
        output_root.mkdir()
        source_payload = _read_bytes_with_long_path(source_path)
        armed = threading.Event()
        injected = threading.Event()
        stop = threading.Event()

        def inject_controlled_fault() -> None:
            armed.set()
            deadline = time.monotonic() + 30.0
            while not stop.is_set() and time.monotonic() < deadline:
                for workspace in physical_runtime_temp.glob("dw-*"):
                    staging_dir = workspace / "staging"
                    materialized = next((workspace / "inputs").glob("input-*.pdf"), None)
                    if materialized is None or not staging_dir.is_dir():
                        continue
                    try:
                        if inject_unresolved_resource:
                            image_dir = staging_dir / f"{materialized.stem}_images"
                            image_dir.mkdir(exist_ok=False)
                            unresolved = image_dir / "unresolved.png"
                            _write_ocr_png(unresolved)
                    except FileExistsError:
                        continue
                    injected.set()
                    return
                time.sleep(0.001)

        injector: threading.Thread | None = None
        if inject_unresolved_resource:
            injector = threading.Thread(target=inject_controlled_fault, name=f"physical-fault-{ordinal}", daemon=True)
            injector.start()
            if not armed.wait(timeout=5):
                raise RuntimeError("packaged_physical_page_fault_injector_not_armed")
        try:
            task_id_value, terminal_value = execute_additional_task(
                plan_request_id=100 + ordinal * 2,
                execute_request_id=101 + ordinal * 2,
                capability_id=capability_id,
                inputs=[
                    {
                        "input_id": f"input.physical.{ordinal}",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": f"inputs/physical-{ordinal}{source_path.suffix}",
                        "locator": {"kind": "local_path", "path": str(source_path)},
                        "media_type": media_type,
                        "size_bytes": len(source_payload),
                        "sha256": hashlib.sha256(source_payload).hexdigest(),
                    }
                ],
                output_root=output_root,
                options={
                    "recognize_text": ocr_enabled,
                    "preserve_resources": keep_images,
                    "ocr_language": "english",
                },
            )
        finally:
            stop.set()
            if injector is not None:
                injector.join(timeout=5)
        if injector is not None and not injected.is_set():
            raise RuntimeError("packaged_physical_page_controlled_fault_not_injected")
        deadline = time.monotonic() + 10.0
        while list(physical_runtime_temp.glob("dw-*")) and time.monotonic() < deadline:
            time.sleep(0.01)
        if list(physical_runtime_temp.glob("dw-*")):
            raise RuntimeError("packaged_physical_page_workspace_residue_after_task")
        return task_id_value, terminal_value, output_root

    physical_results: list[tuple[str, dict[str, Any], Path, bool, bool, int, int, tuple[str, ...] | None]] = []
    for ordinal, (ocr_enabled, keep_images) in enumerate(
        ((False, False), (False, True), (True, False), (True, True)),
        start=1,
    ):
        task_result = execute_physical_task(
            ordinal=ordinal,
            capability_id="convert.pdf.to_markdown",
            source_path=physical_pdf,
            media_type="application/pdf",
            ocr_enabled=ocr_enabled,
            keep_images=keep_images,
            inject_unresolved_resource=keep_images,
        )
        expected_statuses = ("success", "no_text", "recognition_failed", "success") if ocr_enabled else None
        physical_results.append((*task_result, ocr_enabled, keep_images, 4, 5, expected_statuses))
    for ordinal, (ocr_enabled, keep_images) in enumerate(
        ((False, False), (False, True), (True, False), (True, True)),
        start=5,
    ):
        task_result = execute_physical_task(
            ordinal=ordinal,
            capability_id="convert.tiff.to_markdown",
            source_path=physical_tiff,
            media_type="image/tiff",
            ocr_enabled=ocr_enabled,
            keep_images=keep_images,
        )
        expected_statuses = ("success", "no_text", "recognition_failed", "success") if ocr_enabled else None
        physical_results.append((*task_result, ocr_enabled, keep_images, 4, 4, expected_statuses))
    for ordinal, (capability_id, source_path, media_type) in enumerate(
        (
            ("convert.ofd.to_markdown", physical_ofd, "application/vnd.ofd"),
            ("convert.xps.to_markdown", physical_xps, "application/vnd.ms-xpsdocument"),
        ),
        start=9,
    ):
        task_result = execute_physical_task(
            ordinal=ordinal,
            capability_id=capability_id,
            source_path=source_path,
            media_type=media_type,
            ocr_enabled=True,
            keep_images=False,
        )
        physical_results.append((*task_result, True, False, 2, 0, None))

    machine_stdin.close()
    return_code = process.wait(timeout=120)
    trailing_stdout = machine_stdout.read()
    stderr = machine_stderr.read()
    if return_code != 0 or trailing_stdout or stderr:
        raise RuntimeError(f"packaged_machine_protocol_process_failed:{return_code}:{trailing_stdout!r}:{stderr!r}")
    for (
        physical_task_id,
        physical_terminal,
        physical_staging,
        ocr_enabled,
        keep_images,
        page_count,
        resource_count,
        expected_statuses,
    ) in physical_results:
        bundle = _verify_physical_page_bundle(
            terminal=physical_terminal,
            task_id=physical_task_id,
            staging_root=physical_staging,
            page_count=page_count,
            resource_count=resource_count,
            ocr_enabled=ocr_enabled,
            keep_images=keep_images,
            expected_statuses=expected_statuses,
        )
        primary = next(artifact for artifact in bundle["artifacts"] if artifact["kind"] == "document")
        primary_text = _read_text_with_long_path(physical_staging / Path(primary["locator"]))
        if "DOCWEN PHYSICAL PAGE" in primary_text or "DOCWEN TIFF FRAME" in primary_text:
            raise RuntimeError("packaged_physical_page_primary_contains_ocr")
    if terminal.get("method") != "task/completed":
        raise RuntimeError(f"packaged_machine_protocol_terminal_invalid:{terminal}")
    params = terminal.get("params")
    bundle = params.get("bundle") if isinstance(params, dict) else None
    artifacts = bundle.get("artifacts") if isinstance(bundle, dict) else None
    if (
        not isinstance(bundle, dict)
        or bundle.get("task_id") != task_id
        or not isinstance(artifacts, list)
        or len(artifacts) != 1
    ):
        raise RuntimeError(f"packaged_machine_protocol_bundle_invalid:{bundle}")
    artifact = artifacts[0]
    if not isinstance(artifact, dict) or artifact.get("kind") != "document":
        raise RuntimeError(f"packaged_machine_protocol_artifact_invalid:{artifact}")
    locator = artifact.get("locator")
    if not isinstance(locator, str) or "\\" in locator or ".." in locator.split("/"):
        raise RuntimeError(f"packaged_machine_protocol_locator_invalid:{locator}")
    output = staging / Path(locator)
    output_bytes = _read_bytes_with_long_path(output)
    if len(output_bytes) != artifact.get("size_bytes") or hashlib.sha256(output_bytes).hexdigest() != artifact.get(
        "sha256"
    ):
        raise RuntimeError("packaged_machine_protocol_integrity_mismatch")
    with _zipfile_with_long_path(output) as archive:
        if "word/document.xml" not in archive.namelist():
            raise RuntimeError("packaged_machine_protocol_docx_invalid")
        embedded_images = [archive.read(name) for name in archive.namelist() if name.startswith("word/media/")]
    if (
        MACHINE_EXACT_TWO_IMAGE_BYTES not in embedded_images
        or _read_bytes_with_long_path(decoy_image) in embedded_images
    ):
        raise RuntimeError("packaged_machine_protocol_declared_resource_binding_invalid")
    verify_machine_document_semantics_docx(output)
    if semantic_reverse_terminal.get("method") != "task/completed":
        raise RuntimeError(f"packaged_machine_protocol_semantic_reverse_terminal_invalid:{semantic_reverse_terminal}")
    semantic_reverse_params = semantic_reverse_terminal.get("params")
    semantic_reverse_bundle = (
        semantic_reverse_params.get("bundle") if isinstance(semantic_reverse_params, dict) else None
    )
    semantic_reverse_artifacts = (
        semantic_reverse_bundle.get("artifacts") if isinstance(semantic_reverse_bundle, dict) else None
    )
    semantic_reverse_documents = (
        [item for item in semantic_reverse_artifacts if isinstance(item, dict) and item.get("kind") == "document"]
        if isinstance(semantic_reverse_artifacts, list)
        else []
    )
    if (
        not isinstance(semantic_reverse_bundle, dict)
        or semantic_reverse_bundle.get("task_id") != semantic_reverse_task_id
        or len(semantic_reverse_documents) != 1
    ):
        raise RuntimeError(f"packaged_machine_protocol_semantic_reverse_bundle_invalid:{semantic_reverse_bundle}")
    semantic_markdown_locator = semantic_reverse_documents[0].get("locator")
    if (
        not isinstance(semantic_markdown_locator, str)
        or "\\" in semantic_markdown_locator
        or ".." in semantic_markdown_locator.split("/")
    ):
        raise RuntimeError(f"packaged_machine_protocol_semantic_reverse_locator_invalid:{semantic_markdown_locator}")
    semantic_markdown_path = semantic_reverse_staging / Path(semantic_markdown_locator)
    semantic_markdown_bytes = _read_bytes_with_long_path(semantic_markdown_path)
    if len(semantic_markdown_bytes) != semantic_reverse_documents[0].get("size_bytes") or hashlib.sha256(
        semantic_markdown_bytes
    ).hexdigest() != semantic_reverse_documents[0].get("sha256"):
        raise RuntimeError("packaged_machine_protocol_semantic_reverse_integrity_mismatch")
    semantic_markdown = semantic_markdown_bytes.decode("utf-8")
    verify_machine_document_semantics_markdown(semantic_markdown)
    verify_machine_note_domains_markdown(semantic_markdown)
    if ocr_terminal.get("method") != "task/completed":
        raise RuntimeError(f"packaged_machine_protocol_ocr_terminal_invalid:{ocr_terminal}")
    ocr_params = ocr_terminal.get("params")
    ocr_bundle = ocr_params.get("bundle") if isinstance(ocr_params, dict) else None
    ocr_artifacts = ocr_bundle.get("artifacts") if isinstance(ocr_bundle, dict) else None
    ocr_relations = ocr_bundle.get("relations") if isinstance(ocr_bundle, dict) else None
    ocr_manifest_relations = (
        [
            item
            for item in ocr_relations
            if isinstance(item, dict) and item.get("type") == "resource_of" and item.get("role") == "manifest"
        ]
        if isinstance(ocr_relations, list)
        else []
    )
    ocr_manifest_ids = {
        item.get("source_artifact_id")
        for item in ocr_manifest_relations
        if isinstance(item.get("source_artifact_id"), str)
    }
    ocr_manifest_artifacts = (
        [item for item in ocr_artifacts if isinstance(item, dict) and item.get("artifact_id") in ocr_manifest_ids]
        if isinstance(ocr_artifacts, list)
        else []
    )
    ocr_semantic_relations = (
        [item for item in ocr_relations if item not in ocr_manifest_relations]
        if isinstance(ocr_relations, list)
        else []
    )
    if (
        not isinstance(ocr_bundle, dict)
        or ocr_bundle.get("task_id") != ocr_task_id
        or ocr_bundle.get("layout_schema") != "docwen.document_node.v1"
        or not isinstance(ocr_artifacts, list)
        or {item.get("kind") for item in ocr_artifacts if isinstance(item, dict)}
        != {"document", "fragment", "resource"}
        or not isinstance(ocr_relations, list)
        or len(ocr_manifest_relations) != 1
        or len(ocr_manifest_artifacts) != 1
        or ocr_manifest_artifacts[0].get("media_type") != "application/vnd.docwen.document-node+json"
        or ocr_manifest_artifacts[0].get("suggested_name") != "docwen-node.json"
        or [(item.get("type"), item.get("role")) for item in ocr_semantic_relations if isinstance(item, dict)]
        != [
            ("resource_of", "original"),
            ("fragment_of", "ocr_text"),
            ("derived_from", "source"),
        ]
    ):
        raise RuntimeError(f"packaged_machine_protocol_ocr_bundle_invalid:{ocr_bundle}")
    for ocr_artifact in ocr_artifacts:
        if not isinstance(ocr_artifact, dict) or not isinstance(ocr_artifact.get("locator"), str):
            raise RuntimeError(f"packaged_machine_protocol_ocr_artifact_invalid:{ocr_artifact}")
        ocr_output = ocr_staging / Path(ocr_artifact["locator"])
        ocr_output_bytes = _read_bytes_with_long_path(ocr_output)
        if len(ocr_output_bytes) != ocr_artifact.get("size_bytes") or hashlib.sha256(
            ocr_output_bytes
        ).hexdigest() != ocr_artifact.get("sha256"):
            raise RuntimeError("packaged_machine_protocol_ocr_integrity_mismatch")

    def verify_additional_bundle(
        *,
        task_id: str,
        terminal_notification: dict[str, Any],
        output_root: Path,
        expected_kind: str,
        expected_media_type: str,
    ) -> Path:
        if terminal_notification.get("method") != "task/completed":
            raise RuntimeError(f"packaged_machine_protocol_additional_terminal_invalid:{terminal_notification}")
        terminal_params = terminal_notification.get("params")
        task_bundle = terminal_params.get("bundle") if isinstance(terminal_params, dict) else None
        task_artifacts = task_bundle.get("artifacts") if isinstance(task_bundle, dict) else None
        if (
            not isinstance(task_bundle, dict)
            or task_bundle.get("task_id") != task_id
            or not isinstance(task_artifacts, list)
            or len(task_artifacts) != 1
            or not isinstance(task_artifacts[0], dict)
            or task_artifacts[0].get("kind") != expected_kind
            or task_artifacts[0].get("media_type") != expected_media_type
        ):
            raise RuntimeError(f"packaged_machine_protocol_additional_bundle_invalid:{task_bundle}")
        task_artifact = task_artifacts[0]
        task_locator = task_artifact.get("locator")
        if not isinstance(task_locator, str) or "\\" in task_locator or ".." in task_locator.split("/"):
            raise RuntimeError(f"packaged_machine_protocol_additional_locator_invalid:{task_locator}")
        task_output = output_root / Path(task_locator)
        task_output_bytes = _read_bytes_with_long_path(task_output)
        if len(task_output_bytes) != task_artifact.get("size_bytes") or hashlib.sha256(
            task_output_bytes
        ).hexdigest() != task_artifact.get("sha256"):
            raise RuntimeError("packaged_machine_protocol_additional_integrity_mismatch")
        return task_output

    validation_output = verify_additional_bundle(
        task_id=validation_task_id,
        terminal_notification=validation_terminal,
        output_root=validation_staging,
        expected_kind="resource",
        expected_media_type="application/json",
    )
    if not isinstance(json.loads(_read_text_with_long_path(validation_output)), dict):
        raise RuntimeError("packaged_machine_protocol_validation_report_invalid")
    merged_pdf_output = verify_additional_bundle(
        task_id=merge_task_id,
        terminal_notification=merge_terminal,
        output_root=merge_staging,
        expected_kind="document",
        expected_media_type="application/pdf",
    )
    import fitz

    with fitz.open(merged_pdf_output) as merged_document:
        if merged_document.page_count != 2:
            raise RuntimeError("packaged_machine_protocol_merged_pdf_page_count_invalid")
    return output


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Verify packaged DocWenCLI with internal env-var overrides. "
            "DOCWEN_CONFIG_DIR / DOCWEN_LOG_DIR / DOCWEN_LOG_TO_TEMP are internal hooks for CI, "
            "packaging verification, and dev debugging only — not a stable user-facing API."
        )
    )
    parser.add_argument("--binary-dir", required=True, help="Directory containing the packaged DocWenCLI binary.")
    parser.add_argument(
        "--binary-name", default=_default_binary_name(), help="Binary filename inside the package directory."
    )
    parser.add_argument(
        "--ocr-smoke",
        action="store_true",
        help="Also run packaged image->Markdown OCR using the bundled RapidOCR models.",
    )
    parser.add_argument(
        "--proofread-report-smoke",
        action="store_true",
        help="Also verify the packaged Markdown proofread report 2.0 coordinate and empty-result contracts.",
    )
    parser.add_argument(
        "--successful-warning-smoke",
        metavar="INPUT",
        help="Also verify a successful conversion warning in packaged JSON and text modes.",
    )
    parser.add_argument(
        "--successful-warning-action",
        default="gongwen",
        help="Action used by --successful-warning-smoke (default: gongwen).",
    )
    parser.add_argument(
        "--successful-warning-code",
        default="GONGWEN-NEEDS-REVIEW",
        help="Expected warning code for --successful-warning-smoke.",
    )
    parser.add_argument(
        "--successful-warning-message",
        default="",
        help="Optional exact warning message for --successful-warning-smoke.",
    )
    args = parser.parse_args(argv)

    binary_dir = Path(args.binary_dir).resolve()
    binary_path = binary_dir / args.binary_name
    if not binary_path.is_file():
        raise FileNotFoundError(f"packaged_cli_not_found: {binary_path}")
    _verify_resource_layout(binary_dir)

    with _temporary_directory_with_long_path_cleanup(prefix="docwen-packaged-cli-") as temp_dir:
        work_dir = _build_long_path_work_dir(Path(temp_dir))
        work_dir.mkdir(parents=True)
        (work_dir / "config_home").mkdir(parents=True, exist_ok=True)
        (work_dir / "log_home").mkdir(parents=True, exist_ok=True)
        source = work_dir / "示例 数据.xlsx"
        _write_xlsx(source)

        machine_output_file = _run_machine_protocol_smoke(binary_path, work_dir=work_dir)
        _verify_capability_discovery(binary_path, work_dir=work_dir)
        _verify_optimization_discovery(binary_path, work_dir=work_dir)
        template_output_file = _run_template_resource_smoke(binary_path, work_dir=work_dir)
        multiprocessing_egress_report = _run_multiprocessing_egress_boundary_smoke(
            binary_path,
            work_dir=work_dir,
        )

        doctor_process = _run(binary_path, "doctor", "--json", "--quiet", cwd=work_dir)
        _load_verified_doctor_payload(doctor_process)

        pymupdf_layout_output_file = _run_pymupdf_layout_smoke(binary_path, work_dir=work_dir)

        output = work_dir / _LONG_PATH_OUTPUT_NAME
        if len(str(output)) < _LONG_PATH_MINIMUM_LENGTH:
            raise RuntimeError(f"packaged_cli_long_path_fixture_too_short: {output}")
        convert_payload = _load_json_payload(
            _run(
                binary_path,
                "convert",
                str(source),
                "--to",
                "md",
                "--output",
                str(output),
                "--json",
                "--quiet",
                cwd=work_dir,
            ),
            command_name="convert",
        )
        if convert_payload.get("success") is not True:
            raise RuntimeError(f"convert returned success=false: {convert_payload}")
        if convert_payload.get("command") != "convert":
            raise RuntimeError(f"convert payload command mismatch: {convert_payload}")

        output_file = _verify_md_output_file(convert_payload, work_dir=work_dir, command_name="convert")
        converted_text = _read_text_with_long_path(output_file)
        if "name" not in converted_text or "alpha" not in converted_text or "value" not in converted_text:
            raise RuntimeError(f"packaged_convert_output_missing_expected_table_content: {output_file}")
        content_first_output_file = _run_content_first_contract_smoke(binary_path, work_dir=work_dir)
        numbering_added, numbering_removed = _run_numbering_smoke(binary_path, work_dir=work_dir)
        ocr_output_file = _run_optional_ocr_smoke(binary_path, work_dir=work_dir) if args.ocr_smoke else None
        proofread_report_files = (
            _run_optional_proofread_report_smoke(binary_path, work_dir=work_dir)
            if args.proofread_report_smoke
            else None
        )
        warning_output_file = None
        if args.successful_warning_smoke:
            warning_input = Path(args.successful_warning_smoke).resolve()
            if not warning_input.is_file():
                raise FileNotFoundError(f"packaged_successful_warning_input_missing: {warning_input}")
            warning_output_file = _run_optional_successful_warning_smoke(
                binary_path,
                work_dir=work_dir,
                input_path=warning_input,
                action=args.successful_warning_action,
                expected_code=args.successful_warning_code,
                expected_message=args.successful_warning_message,
            )

        log_dir = work_dir / "log_home" / "logs"
        log_files = list(log_dir.glob("*.log"))
        if not log_files:
            raise RuntimeError(f"packaged_cli_log_missing: {log_dir}")

        message = f"packaged_cli_smoke_ok: {binary_path.name} -> {output_file.name}"
        message += f"; machine-v1 -> {machine_output_file.name}"
        message += f"; egress-boundary -> {multiprocessing_egress_report.name}"
        message += f"; pymupdf-layout -> {pymupdf_layout_output_file.name}"
        message += f"; content-first -> {content_first_output_file.name}"
        message += f"; template-id -> {template_output_file.name}"
        message += f"; numbering -> {numbering_added.name} / {numbering_removed.name}"
        if ocr_output_file is not None:
            message += f"; ocr -> {ocr_output_file.name}"
        if proofread_report_files is not None:
            message += f"; proofread-report -> {proofread_report_files[0].name} / {proofread_report_files[1].name}"
        if warning_output_file is not None:
            message += f"; warning {args.successful_warning_code} -> {warning_output_file.name}"
        print(message)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
