"""Real ``docwen serve --stdio`` document round-trip source-tree E2E."""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any, BinaryIO, cast

import pytest
from openpyxl import Workbook, load_workbook
from PIL import Image, ImageDraw, ImageFont
from scripts.release.verify_packaged_cli import (
    MACHINE_DOCUMENT_SEMANTICS_FIXTURE,
    MACHINE_DOCUMENT_SEMANTICS_LIMITATIONS,
    MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT,
    MACHINE_EXACT_TWO_NUMBERING_PLAN,
    MACHINE_NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    MACHINE_RESOLVED_DOCUMENT_LIMITATIONS,
    MACHINE_RESOLVED_DOCUMENT_MEDIA_TYPE,
    verify_machine_document_semantics_docx,
    verify_machine_document_semantics_markdown,
    verify_machine_note_domains_markdown,
)
from tools.validate_contracts import validate_trace

from docwen_cli.machine.contracts import MachineContractValidator
from docwen_cli.machine.framing import FrameWriter, read_frame


def _request(request_id: int, method: str, params: dict[str, Any]) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "method": method, "params": params}


def _write_ocr_png(path: Path) -> None:
    image = Image.new("RGB", (900, 240), "white")
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 56)
    except OSError:
        font = ImageFont.load_default()
    draw.text((36, 72), "HELLO DOCWEN OCR", fill="black", font=font)
    image.save(path)


def _file_inventory(path: Path, *, relative_to: Path) -> dict[str, str | int]:
    payload = path.read_bytes()
    return {
        "path": path.relative_to(relative_to).as_posix(),
        "size_bytes": len(payload),
        "sha256": hashlib.sha256(payload).hexdigest(),
    }


def _directory_inventory(root: Path) -> tuple[dict[str, str | int], ...]:
    return tuple(
        _file_inventory(path, relative_to=root)
        for path in sorted(root.rglob("*"), key=lambda item: item.as_posix())
        if path.is_file()
    )


def _create_directory_link(link: Path, target: Path) -> None:
    if os.name == "nt":
        completed = subprocess.run(
            ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target.resolve(strict=True))],
            capture_output=True,
            text=True,
            errors="replace",
            check=False,
        )
        if completed.returncode != 0:
            if os.path.lexists(link):
                os.rmdir(link)
            pytest.skip(f"NTFS junction creation is unavailable: {completed.stderr or completed.stdout}")
        return
    try:
        link.symlink_to(target, target_is_directory=True)
    except (NotImplementedError, OSError) as exc:
        if os.path.lexists(link):
            link.unlink()
        pytest.skip(f"directory symlink creation is unavailable: {exc}")


def _remove_directory_link(link: Path) -> None:
    """Remove only the link node without traversing its directory target."""

    if os.name == "nt":
        os.rmdir(link)
    else:
        link.unlink()


def _exercise_auxiliary_capability_matrix(
    *,
    tmp_path: Path,
    workbook_source: Path,
    pdf_source: Path,
    tiff_source: Path,
    ocr_source: Path,
    tables_source: Path,
    source: Path,
    source_bytes: bytes,
    process_stdout: BinaryIO,
    exchange: Any,
    trace: list[dict[str, Any]],
) -> None:
    workbook_staging = tmp_path / "workbook-staging"
    workbook_staging.mkdir()
    workbook_bytes = workbook_source.read_bytes()
    workbook_plan = exchange(
        _request(
            7,
            "task/plan",
            {
                "capability_id": "convert.xlsx.to_csv",
                "inputs": [
                    {
                        "input_id": "input.3",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": "inputs/workbook.xlsx",
                        "locator": {"kind": "local_path", "path": str(workbook_source)},
                        "media_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "size_bytes": len(workbook_bytes),
                        "sha256": hashlib.sha256(workbook_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(workbook_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        )
    )
    exchange(_request(8, "task/execute", {"plan_id": workbook_plan["result"]["plan_id"]}))
    while True:
        workbook_terminal = read_frame(process_stdout)
        assert workbook_terminal is not None
        trace.append(workbook_terminal)
        if workbook_terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            break
    assert workbook_terminal["method"] == "task/completed"
    workbook_bundle = workbook_terminal["params"]["bundle"]
    assert [artifact["kind"] for artifact in workbook_bundle["artifacts"]] == ["resource", "resource"]
    assert [(entry["role"], entry["ordinal"], entry["preferred"]) for entry in workbook_bundle["entries"]] == [
        ("worksheet", 0, True),
        ("worksheet", 1, False),
    ]
    assert workbook_bundle["relations"] == []
    for artifact in workbook_bundle["artifacts"]:
        path = workbook_staging / Path(artifact["locator"])
        assert path.stat().st_size == artifact["size_bytes"]
        assert hashlib.sha256(path.read_bytes()).hexdigest() == artifact["sha256"]

    pdf_staging = tmp_path / "pdf-staging"
    pdf_staging.mkdir()
    pdf_bytes = pdf_source.read_bytes()
    pdf_plan = exchange(
        _request(
            9,
            "task/plan",
            {
                "capability_id": "render.pdf.to_png",
                "inputs": [
                    {
                        "input_id": "input.4",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": "inputs/pages.pdf",
                        "locator": {"kind": "local_path", "path": str(pdf_source)},
                        "media_type": "application/pdf",
                        "size_bytes": len(pdf_bytes),
                        "sha256": hashlib.sha256(pdf_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(pdf_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        )
    )
    exchange(_request(10, "task/execute", {"plan_id": pdf_plan["result"]["plan_id"]}))
    while True:
        pdf_terminal = read_frame(process_stdout)
        assert pdf_terminal is not None
        trace.append(pdf_terminal)
        if pdf_terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            break
    assert pdf_terminal["method"] == "task/completed"
    pdf_bundle = pdf_terminal["params"]["bundle"]
    assert [artifact["kind"] for artifact in pdf_bundle["artifacts"]] == ["resource", "resource"]
    assert [(entry["role"], entry["ordinal"], entry["preferred"]) for entry in pdf_bundle["entries"]] == [
        ("image", 0, True),
        ("image", 1, False),
    ]
    for artifact in pdf_bundle["artifacts"]:
        path = pdf_staging / Path(artifact["locator"])
        assert path.stat().st_size == artifact["size_bytes"]
        assert hashlib.sha256(path.read_bytes()).hexdigest() == artifact["sha256"]

    physical_pdf_staging = tmp_path / "physical-pdf-staging"
    physical_pdf_staging.mkdir()
    physical_pdf_plan = exchange(
        _request(
            101,
            "task/plan",
            {
                "capability_id": "convert.pdf.to_markdown",
                "inputs": [
                    {
                        "input_id": "input.physical.pdf",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": "inputs/physical-pages.pdf",
                        "locator": {"kind": "local_path", "path": str(pdf_source)},
                        "media_type": "application/pdf",
                        "size_bytes": len(pdf_bytes),
                        "sha256": hashlib.sha256(pdf_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(physical_pdf_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {"recognize_text": True, "preserve_resources": False},
            },
        )
    )
    assert physical_pdf_plan["result"]["output_shape"]["relation_payloads"] == [
        "page_fragment",
        "page_resource",
    ]
    exchange(_request(102, "task/execute", {"plan_id": physical_pdf_plan["result"]["plan_id"]}))
    while True:
        physical_pdf_terminal = read_frame(process_stdout)
        assert physical_pdf_terminal is not None
        trace.append(physical_pdf_terminal)
        if physical_pdf_terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            break
    assert physical_pdf_terminal["method"] == "task/completed", physical_pdf_terminal
    physical_pdf_bundle = physical_pdf_terminal["params"]["bundle"]
    assert [artifact["kind"] for artifact in physical_pdf_bundle["artifacts"]].count("document") == 1
    assert [artifact["kind"] for artifact in physical_pdf_bundle["artifacts"]].count("fragment") == 2
    assert any(
        artifact["media_type"] == "application/vnd.docwen.document-node+json"
        for artifact in physical_pdf_bundle["artifacts"]
    )
    page_relations = [
        relation
        for relation in physical_pdf_bundle["relations"]
        if relation["type"] == "fragment_of" and relation["role"] == "ocr_page"
    ]
    assert [relation["page_fragment"]["page_index"] for relation in page_relations] == [1, 2]
    assert [relation["page_fragment"]["page_count"] for relation in page_relations] == [2, 2]
    assert [relation["ordinal"] for relation in page_relations] == [0, 1]
    for artifact in physical_pdf_bundle["artifacts"]:
        path = physical_pdf_staging / Path(artifact["locator"])
        assert path.stat().st_size == artifact["size_bytes"]
        assert hashlib.sha256(path.read_bytes()).hexdigest() == artifact["sha256"]

    physical_tiff_staging = tmp_path / "physical-tiff-staging"
    physical_tiff_staging.mkdir()
    physical_tiff_bytes = tiff_source.read_bytes()
    physical_tiff_plan = exchange(
        _request(
            103,
            "task/plan",
            {
                "capability_id": "convert.tiff.to_markdown",
                "inputs": [
                    {
                        "input_id": "input.physical.tiff",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": "inputs/physical-frames.tiff",
                        "locator": {"kind": "local_path", "path": str(tiff_source)},
                        "media_type": "image/tiff",
                        "size_bytes": len(physical_tiff_bytes),
                        "sha256": hashlib.sha256(physical_tiff_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(physical_tiff_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {"recognize_text": False, "preserve_resources": True},
            },
        )
    )
    exchange(_request(104, "task/execute", {"plan_id": physical_tiff_plan["result"]["plan_id"]}))
    while True:
        physical_tiff_terminal = read_frame(process_stdout)
        assert physical_tiff_terminal is not None
        trace.append(physical_tiff_terminal)
        if physical_tiff_terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            break
    assert physical_tiff_terminal["method"] == "task/completed", physical_tiff_terminal
    physical_tiff_bundle = physical_tiff_terminal["params"]["bundle"]
    assert [artifact["kind"] for artifact in physical_tiff_bundle["artifacts"]].count("document") == 1
    assert [artifact["kind"] for artifact in physical_tiff_bundle["artifacts"]].count("resource") == 3
    image_relations = [
        relation
        for relation in physical_tiff_bundle["relations"]
        if relation["type"] == "resource_of" and relation["role"] == "image"
    ]
    assert [relation["page_resource"]["source_page"] for relation in image_relations] == [1, 2]

    split_staging = tmp_path / "split-staging"
    split_staging.mkdir()
    split_plan = exchange(
        _request(
            11,
            "task/plan",
            {
                "capability_id": "split.pdf.every_page",
                "inputs": [
                    {
                        "input_id": "input.5",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": "inputs/split-pages.pdf",
                        "locator": {"kind": "local_path", "path": str(pdf_source)},
                        "media_type": "application/pdf",
                        "size_bytes": len(pdf_bytes),
                        "sha256": hashlib.sha256(pdf_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(split_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        )
    )
    assert split_plan["result"]["effective_options"] == {"split_mode": "every_page"}
    exchange(_request(12, "task/execute", {"plan_id": split_plan["result"]["plan_id"]}))
    while True:
        split_terminal = read_frame(process_stdout)
        assert split_terminal is not None
        trace.append(split_terminal)
        if split_terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            break
    assert split_terminal["method"] == "task/completed"
    split_bundle = split_terminal["params"]["bundle"]
    assert [artifact["kind"] for artifact in split_bundle["artifacts"]] == ["document", "document"]
    assert [(entry["role"], entry["ordinal"], entry["preferred"]) for entry in split_bundle["entries"]] == [
        ("section", 0, True),
        ("section", 1, False),
    ]
    assert split_bundle["relations"] == []

    ocr_staging = tmp_path / "ocr-staging"
    ocr_staging.mkdir()
    ocr_bytes = ocr_source.read_bytes()
    ocr_plan = exchange(
        _request(
            13,
            "task/plan",
            {
                "capability_id": "convert.png.to_ocr_markdown",
                "inputs": [
                    {
                        "input_id": "input.6",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": "inputs/ocr-source.png",
                        "locator": {"kind": "local_path", "path": str(ocr_source)},
                        "media_type": "image/png",
                        "size_bytes": len(ocr_bytes),
                        "sha256": hashlib.sha256(ocr_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(ocr_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        )
    )
    assert ocr_plan["result"]["effective_options"] == {
        "image_mode": "file",
        "to_md_keep_images": True,
        "to_md_enable_ocr": True,
        "ocr_placement": "image_md",
    }
    exchange(_request(14, "task/execute", {"plan_id": ocr_plan["result"]["plan_id"]}))
    while True:
        ocr_terminal = read_frame(process_stdout)
        assert ocr_terminal is not None
        trace.append(ocr_terminal)
        if ocr_terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            break
    assert ocr_terminal["method"] == "task/completed"
    ocr_bundle = ocr_terminal["params"]["bundle"]
    ocr_artifacts = {artifact["kind"]: artifact for artifact in ocr_bundle["artifacts"]}
    assert set(ocr_artifacts) == {"document", "fragment", "resource"}
    fragment_path = ocr_staging / Path(ocr_artifacts["fragment"]["locator"])
    assert "DOCWEN" in fragment_path.read_text(encoding="utf-8").upper()
    assert [(relation["type"], relation["role"]) for relation in ocr_bundle["relations"]] == [
        ("resource_of", "original"),
        ("fragment_of", "ocr_text"),
        ("derived_from", "source"),
        ("resource_of", "manifest"),
    ]

    table_staging = tmp_path / "table-staging"
    table_staging.mkdir()
    table_bytes = tables_source.read_bytes()
    table_plan = exchange(
        _request(
            15,
            "task/plan",
            {
                "capability_id": "convert.markdown_tables.to_csv",
                "inputs": [
                    {
                        "input_id": "input.7",
                        "kind": "document",
                        "role": "source",
                        "logical_path": "inputs/tables.md",
                        "locator": {"kind": "local_path", "path": str(tables_source)},
                        "media_type": "text/markdown",
                        "size_bytes": len(table_bytes),
                        "sha256": hashlib.sha256(table_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(table_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        )
    )
    exchange(_request(16, "task/execute", {"plan_id": table_plan["result"]["plan_id"]}))
    while True:
        table_terminal = read_frame(process_stdout)
        assert table_terminal is not None
        trace.append(table_terminal)
        if table_terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            break
    assert table_terminal["method"] == "task/completed"
    table_bundle = table_terminal["params"]["bundle"]
    assert [artifact["kind"] for artifact in table_bundle["artifacts"]] == ["resource", "resource"]
    assert [(entry["role"], entry["ordinal"]) for entry in table_bundle["entries"]] == [
        ("supplementary", 0),
        ("supplementary", 1),
    ]

    frame_staging = tmp_path / "frame-staging"
    frame_staging.mkdir()
    tiff_bytes = tiff_source.read_bytes()
    frame_plan = exchange(
        _request(
            17,
            "task/plan",
            {
                "capability_id": "convert.tiff_frames.to_png",
                "inputs": [
                    {
                        "input_id": "input.8",
                        "kind": "resource",
                        "role": "source",
                        "logical_path": "inputs/frames.tiff",
                        "locator": {"kind": "local_path", "path": str(tiff_source)},
                        "media_type": "image/tiff",
                        "size_bytes": len(tiff_bytes),
                        "sha256": hashlib.sha256(tiff_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(frame_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        )
    )
    exchange(_request(18, "task/execute", {"plan_id": frame_plan["result"]["plan_id"]}))
    while True:
        frame_terminal = read_frame(process_stdout)
        assert frame_terminal is not None
        trace.append(frame_terminal)
        if frame_terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            break
    assert frame_terminal["method"] == "task/completed"
    frame_bundle = frame_terminal["params"]["bundle"]
    assert [artifact["kind"] for artifact in frame_bundle["artifacts"]] == ["resource", "resource"]
    assert [(entry["role"], entry["ordinal"]) for entry in frame_bundle["entries"]] == [
        ("image", 0),
        ("image", 1),
    ]

    health = exchange(_request(19, "health/check", {}))
    assert health["result"]["all_ok"] is True
    inspected = exchange(
        _request(
            20,
            "file/inspect",
            {
                "input": {
                    "input_id": "input.inspect",
                    "kind": "document",
                    "role": "source",
                    "logical_path": "inputs/inspect.md",
                    "locator": {"kind": "local_path", "path": str(source)},
                    "media_type": "text/markdown",
                    "size_bytes": len(source_bytes),
                    "sha256": hashlib.sha256(source_bytes).hexdigest(),
                }
            },
        )
    )
    assert inspected["result"]["detected_format"] in {"md", "markdown"}
    assert inspected["result"]["size_bytes"] == len(source_bytes)
    assert inspected["result"]["content_sha256"] == hashlib.sha256(source_bytes).hexdigest()
    resources = exchange(_request(21, "resource/list", {"kind": "numbering-schemes", "locale": "en_US"}))
    assert resources["result"]["kind"] == "numbering-schemes"
    assert resources["result"]["resources"]


__all__ = (
    "MACHINE_DOCUMENT_SEMANTICS_FIXTURE",
    "MACHINE_DOCUMENT_SEMANTICS_LIMITATIONS",
    "MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT",
    "MACHINE_EXACT_TWO_NUMBERING_PLAN",
    "MACHINE_NUMBERING_EXPORT_PLAN_MEDIA_TYPE",
    "MACHINE_RESOLVED_DOCUMENT_LIMITATIONS",
    "MACHINE_RESOLVED_DOCUMENT_MEDIA_TYPE",
    "Any",
    "BinaryIO",
    "FrameWriter",
    "Image",
    "MachineContractValidator",
    "Path",
    "Workbook",
    "_create_directory_link",
    "_directory_inventory",
    "_exercise_auxiliary_capability_matrix",
    "_file_inventory",
    "_remove_directory_link",
    "_request",
    "_write_ocr_png",
    "cast",
    "hashlib",
    "json",
    "load_workbook",
    "os",
    "pytest",
    "read_frame",
    "subprocess",
    "sys",
    "validate_trace",
    "verify_machine_document_semantics_docx",
    "verify_machine_document_semantics_markdown",
    "verify_machine_note_domains_markdown",
    "zipfile",
)
