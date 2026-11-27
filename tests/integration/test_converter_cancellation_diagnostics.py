"""Cancelling active conversions reports one cancellation and leaves sources intact."""

from pathlib import Path
from zipfile import ZipFile

import fitz
import pytest
from docx import Document
from PIL import Image

from docwen_bundle.runtime_factory import create_runtime_port
from docwen_core.models import ConversionRequest, FileRef, OutputPolicy

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def _write_source(path: Path) -> None:
    if path.suffix == ".pdf":
        with fitz.open() as document:
            document.new_page().insert_text((30, 50), "First page")
            document.new_page().insert_text((30, 50), "Second page")
            document.save(path)
    elif path.suffix == ".docx":
        document = Document()
        document.add_paragraph("Text to proofread")
        document.save(path)
    elif path.suffix == ".png":
        with Image.new("RGB", (8, 8), "white") as image:
            image.save(path)
    elif path.suffix == ".ofd":
        with ZipFile(path, "w") as package:
            package.writestr("OFD.xml", "<OFD><DocBody><DocRoot>body/Document.xml</DocRoot></DocBody></OFD>")
            package.writestr("body/Document.xml", "<Document/>")
            package.writestr("body/InvoiceData.xml", "<Invoice><InvoiceNumber>00000028</InvoiceNumber></Invoice>")


@pytest.mark.parametrize(
    ("source_format", "category", "target", "action", "options", "cancel_at"),
    [
        ("pdf", "layout", "pdf", "merge_pdfs", {}, "Adding file 1/2"),
        ("pdf", "layout", "pdf", "split_pdf", {"split_mode": "every_page"}, "Starting PDF split"),
        ("pdf", "layout", "md", "", {}, "Starting layout to Markdown conversion"),
        ("docx", "document", "docx", "validate", {"enable_symbol_pairing": True}, "Loading DOCX document"),
        ("png", "image", "md", "", {"to_md_enable_ocr": False}, "Starting image to Markdown conversion"),
        ("ofd", "layout", "md", "invoice_cn", {}, "Extracting OFD invoice data"),
    ],
    ids=["merge-pdf", "split-pdf", "pdf-to-md", "proofread-docx", "image-to-md", "invoice-to-md"],
)
def test_cancel_during_conversion_has_no_spurious_failure_diagnostic(
    tmp_path, source_format, category, target, action, options, cancel_at
):
    sources = [tmp_path / f"first.{source_format}"]
    if action == "merge_pdfs":
        sources.append(tmp_path / "second.pdf")
    for path in sources:
        _write_source(path)
    original = {path: path.read_bytes() for path in sources}
    output = tmp_path / "published"
    request = ConversionRequest(
        request_id="cancel-in-converter",
        input_refs=[FileRef(path=str(path), format=source_format, category=category) for path in sources],
        target_format=target,
        action_name=action,
        options=options,
        output_policy=OutputPolicy(output_dir=str(output)),
    )
    reached = []

    def cancel_on_progress(event):
        if event.event_type == "task_progress" and event.payload.get("message") == cancel_at:
            reached.append(event.task_id)
            runtime.cancel(event.task_id)

    runtime = create_runtime_port(event_callback=cancel_on_progress)
    result = runtime.execute(request)
    assert reached == [request.request_id]
    assert not result.success and result.error.error_type == "cancelled"
    terminal_types = {"task_completed", "task_failed", "task_cancelled"}
    assert [e.event_type for e in runtime.collected_events if e.event_type in terminal_types] == ["task_cancelled"]
    assert [diagnostic.code for diagnostic in result.diagnostics if diagnostic.level == "error"] == []
    assert result.artifacts == [] and not list(output.rglob("*"))
    assert {path: path.read_bytes() for path in sources} == original
