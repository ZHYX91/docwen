"""Frozen HTML/plain inputs through Store, Runtime and final DOCX Bundle.

The captures are synthesized bytes, not observations of an installed clipboard
provider. OOXML structure and exact authored text are checked independently.
"""

from __future__ import annotations

import hashlib
import xml.etree.ElementTree as ET
from pathlib import Path
from types import SimpleNamespace
from typing import TYPE_CHECKING, Any, cast
from zipfile import ZipFile

import pytest

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_core.detection import inspect_utf8_markdown_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.clipboard_document import clipboard_document_to_bytes
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_gui.clipboard_capture import FrozenClipboardCapture
from docwen_gui.clipboard_inputs import ClipboardInputStore
from docwen_gui.clipboard_rich_document import project_frozen_rich_document
from docwen_gui.execution_requests import ExecutionRequestBuilder
from docwen_gui.path_identity import normalize_path
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
from docwen_runtime.templates import TemplateRegistry
from tests.integration.test_structured_resource_group_runtime import _builder, _source_ref

if TYPE_CHECKING:
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{urn:docwen:clipboard-table-associations:v1}"
OVERSIZED_HTML = b'<table><tr><td rowspan="2" colspan="65536">A</td></tr><tr></tr></table>'
LEGAL_HTML = (
    b"<p>Before</p><table><thead><tr>"
    b'<th colspan="2" id="heading" scope="colgroup">Header</th></tr></thead>'
    b'<tbody><tr><th rowspan="0" scope="rowgroup">Group</th><td headers="heading">'
    b"<p>before</p><table><tr><td>inner</td></tr></table><p>after</p></td></tr>"
    b'<tr><td headers="heading">Tail</td></tr></tbody>'
    b"<tfoot><tr><td>FootA</td><td>FootB</td></tr></tfoot></table><p>After</p>"
)

# With no row-header assertion, the unchanged colspan header is natively representable.
NATIVE_HEADER_HTML = LEGAL_HTML.replace(b'<th rowspan="0" scope="rowgroup">Group</th>', b'<td rowspan="0">Group</td>')


def _word_bundle(runtime: Any, request: ConversionRequest, output: Path) -> tuple[Path, ET.Element]:
    result = runtime.execute(request)
    assert result.success, result.error
    bundle = ArtifactBundleCommitter().commit(
        task_id=request.request_id,
        staging_root=str(output),
        draft=build_bundle_draft(
            profile="single_document",
            output_media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            artifacts=result.artifacts,
        ),
    )
    primary_id = next(entry.artifact_id for entry in bundle.entries if entry.preferred)
    primary = next(item for item in bundle.artifacts if item.artifact_id == primary_id)
    path = output / primary.locator
    data = path.read_bytes()
    assert len(data) == primary.size_bytes
    assert hashlib.sha256(data).hexdigest() == primary.sha256
    with ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    body = root.find(f"{W}body")
    assert body is not None
    return path, body


def _text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.iter(f"{W}t"))


def test_oversized_html_keeps_complete_plain_snapshot_and_published_body(
    tmp_path: Path, round_trip_runtime: Any
) -> None:
    html_path = tmp_path / "oversized.html"
    html_path.write_bytes(OVERSIZED_HTML)
    plain = "Before preserved\n\nA\n\nAfter preserved\n"
    plain_path = tmp_path / "captured.txt"
    plain_path.write_bytes(plain.encode("utf-8"))
    capture = FrozenClipboardCapture(plain_path.read_text(), html_path.read_bytes(), (), None)
    decision = project_frozen_rich_document(capture)
    assert decision.handled and decision.plain_fallback and decision.projection is None
    assert capture.plain_text is not None
    store = ClipboardInputStore(tmp_path / "managed")
    try:
        snapshot = store.create(capture.plain_text, display_name_template="Clipboard {index}.md")
        store.sync_visible([snapshot.path])
        inspection = inspect_utf8_markdown_snapshot(snapshot.path)
        assert Path(snapshot.path).read_bytes() == plain.encode("utf-8")
        template = next(
            item.id
            for item in TemplateRegistry.default().list_templates("docx")
            if item.name == "English General Template"
        )
        output = tmp_path / "published"
        source_ref = FileRef(
            path=snapshot.path,
            format="markdown",
            category="markdown",
            size_bytes=snapshot.size_bytes,
            metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
        )
        builder = ExecutionRequestBuilder(
            cast("MainWindowViewModel", SimpleNamespace(files=[source_ref], controller=None)),
            cast("BatchListViewModel", SimpleNamespace(get_file_entry=lambda _path: None)),
            file_contexts=lambda: {normalize_path(snapshot.path): ("markdown", "markdown")},
            selected_template=lambda: ("docx", template),
            source_label=lambda path: snapshot.display_name if store.is_snapshot(path) else None,
            synthetic_input=store.is_snapshot,
        )
        request, _ = builder.single(
            file_path=snapshot.path,
            target_format="docx",
            action_name="",
            options={"remove_numbering": False},
            output_policy=OutputPolicy(output_dir=str(output)),
        )
        assert request.source_stem == "Clipboard 1"
        _, body = _word_bundle(round_trip_runtime, request, output)
        paragraphs = [_text(node) for node in body.findall(f"{W}p") if _text(node)]
        # The selected template owns this one title paragraph; the entire authored body follows.
        assert paragraphs == ["Clipboard 1", "Before preserved", "A", "After preserved"]
        assert body.findall(f"{W}tbl") == []
        assert Path(snapshot.path).read_bytes() == plain.encode("utf-8")
        assert html_path.read_bytes() == OVERSIZED_HTML
        assert plain_path.read_bytes() == plain.encode("utf-8")
    finally:
        store.close()


@pytest.mark.parametrize(
    ("html", "header_rows", "group_scope"),
    [(LEGAL_HTML, 0, "rowgroup"), (NATIVE_HEADER_HTML, 1, "")],
    ids=["html-association-not-native-prefix", "representable-native-header"],
)
def test_legal_row_group_merge_nested_order_and_headers_survive_final_package(
    tmp_path: Path, round_trip_runtime: Any, html: bytes, header_rows: int, group_scope: str
) -> None:
    html_path = tmp_path / "legal.html"
    html_path.write_bytes(html)
    capture = FrozenClipboardCapture(None, html_path.read_bytes(), (), None)
    decision = project_frozen_rich_document(capture)
    assert decision.handled and not decision.plain_fallback
    projection = decision.projection
    assert projection is not None and projection.html_only
    payload = clipboard_document_to_bytes(projection.document)
    (tmp_path / "source.dwclip").write_bytes(payload)
    store = ClipboardInputStore(tmp_path / "managed")
    try:
        snapshot = store.create_bundle(
            payload,
            display_name_template="Clipboard {index}.dwclip",
            preview="Legal merge",
            resources=projection.resources,
        )
        store.sync_visible([snapshot.main.path])
        output = tmp_path / "published"
        request, _ = _builder(store, [_source_ref(snapshot)]).single(
            file_path=snapshot.main.path,
            target_format="docx",
            action_name="",
            options={},
            output_policy=OutputPolicy(output_dir=str(output)),
        )
        path, body = _word_bundle(round_trip_runtime, request, output)
        [table] = body.findall(f"{W}tbl")
        rows = table.findall(f"{W}tr")
        assert len(rows) == 4
        assert [node.text for node in table.iter(f"{W}t")] == [
            "Header",
            "Group",
            "before",
            "inner",
            "after",
            "Tail",
            "FootA",
            "FootB",
        ]
        header_cells = rows[0].findall(f"{W}tc")
        assert len(header_cells) == 1
        span = header_cells[0].find(f"{W}tcPr/{W}gridSpan")
        assert span is not None and span.get(f"{W}val") == "2"
        native_header = rows[0].find(f"{W}trPr/{W}tblHeader")
        assert (native_header is not None) == bool(header_rows)
        assert all(row.find(f"{W}trPr/{W}tblHeader") is None for row in rows[1:])
        first = rows[1].findall(f"{W}tc")
        continuation = rows[2].findall(f"{W}tc")[0]
        restart = first[0].find(f"{W}tcPr/{W}vMerge")
        continued = continuation.find(f"{W}tcPr/{W}vMerge")
        assert restart is not None and restart.get(f"{W}val") == "restart"
        assert continued is not None and continued.get(f"{W}val") in {None, "continue"}
        assert _text(continuation) == ""
        nested_blocks = [
            (node.tag, _text(node)) for node in first[1] if node.tag in {f"{W}p", f"{W}tbl"} and _text(node)
        ]
        assert nested_blocks == [(f"{W}p", "before"), (f"{W}tbl", "inner"), (f"{W}p", "after")]
        assert [_text(node) for node in body.findall(f"{W}p") if _text(node)] == ["Before", "After"]
        with ZipFile(path) as package:
            maps = [
                root
                for name in package.namelist()
                if name.startswith("customXml/item") and name.endswith(".xml")
                for root in [ET.fromstring(package.read(name))]
                if root.tag == f"{A}clipboardTableAssociations"
            ]
        [association] = maps
        outer = association.find(f"{A}table")
        assert outer is not None
        assert outer.get("index") == "0" and outer.get("header_rows") == str(header_rows)
        assert outer.get("header_columns") == "0"
        cells = outer.findall(f"{A}cell")
        assert cells[0].get("html_id") == "heading" and cells[0].get("scope") == "colgroup"
        assert cells[1].get("scope") == group_scope and cells[1].get("row_span") == "2"
        assert cells[1].get("header") == ("1" if group_scope else "0")
        assert [[node.get("ref") for node in cell.findall(f"{A}header")] for cell in cells] == [
            [],
            [],
            ["heading"],
            ["heading"],
            [],
            [],
        ]
        assert Path(snapshot.main.path).read_bytes() == payload
        assert html_path.read_bytes() == html
    finally:
        store.close()
