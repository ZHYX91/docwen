"""Focused tests split from test_machine_stdio_e2e.py."""

from __future__ import annotations

from docwen_core.round_trip_sidecar import ROUND_TRIP_SIDECAR_MEDIA_TYPE, read_round_trip_sidecar

from ._machine_stdio_e2e_support import (
    MACHINE_DOCUMENT_SEMANTICS_FIXTURE,
    MACHINE_DOCUMENT_SEMANTICS_LIMITATIONS,
    MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT,
    MACHINE_EXACT_TWO_NUMBERING_PLAN,
    MACHINE_NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    MACHINE_RESOLVED_DOCUMENT_LIMITATIONS,
    MACHINE_RESOLVED_DOCUMENT_MEDIA_TYPE,
    Any,
    BinaryIO,
    FrameWriter,
    Image,
    MachineContractValidator,
    Path,
    Workbook,
    _exercise_auxiliary_capability_matrix,
    _request,
    _write_ocr_png,
    cast,
    hashlib,
    json,
    pytest,
    read_frame,
    subprocess,
    sys,
    validate_trace,
    verify_machine_document_semantics_docx,
    verify_machine_document_semantics_markdown,
    verify_machine_note_domains_markdown,
    zipfile,
)


@pytest.mark.e2e
def test_real_stdio_process_emits_integrity_pinned_docx_bundle(tmp_path: Path) -> None:
    source_dir = tmp_path / "source-physical"
    source_dir.mkdir()
    source = source_dir / "source.md"
    decoy_dir = source_dir / "assets"
    decoy_dir.mkdir()
    (decoy_dir / "机器协议-语义.png").write_bytes(b"physical sibling decoy")
    source.write_text(MACHINE_DOCUMENT_SEMANTICS_FIXTURE, encoding="utf-8")
    source_bytes = source.read_bytes()
    neutral_document = tmp_path / "neutral-document.json"
    neutral_document.write_text(
        json.dumps(MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    numbering_plan = tmp_path / "numbering-export-plan.json"
    numbering_plan.write_text(
        json.dumps(MACHINE_EXACT_TWO_NUMBERING_PLAN, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    neutral_bytes = neutral_document.read_bytes()
    numbering_plan_bytes = numbering_plan.read_bytes()
    workbook_source = tmp_path / "workbook.xlsx"
    workbook = Workbook()
    active_sheet = workbook.active
    assert active_sheet is not None
    active_sheet.title = "First"
    active_sheet.append(["name", "value"])
    active_sheet.append(["alpha", 1])
    second = workbook.create_sheet("Second")
    second.append(["beta", 2])
    workbook.save(workbook_source)
    workbook.close()
    pdf_source = tmp_path / "pages.pdf"
    import fitz

    pdf = fitz.open()
    for page_number in range(1, 3):
        page = pdf.new_page()
        page.insert_text((72, 72), f"Page {page_number}")
    pdf.save(pdf_source)
    pdf.close()
    ocr_source = tmp_path / "ocr-source.png"
    _write_ocr_png(ocr_source)
    tables_source = tmp_path / "tables.md"
    tables_source.write_text(
        "# Tables\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n| C | D |\n|---|---|\n| 3 | 4 |\n",
        encoding="utf-8",
    )
    tiff_source = tmp_path / "frames.tiff"
    tiff_frames = [Image.new("RGB", (8, 8), color) for color in ((255, 0, 0), (0, 0, 255))]
    tiff_frames[0].save(tiff_source, save_all=True, append_images=tiff_frames[1:], format="TIFF")
    for frame in tiff_frames:
        frame.close()
    staging = tmp_path / "staging"
    staging.mkdir()
    process = subprocess.Popen(
        [
            sys.executable,
            "-c",
            "from docwen_bundle.cli_entry import main; raise SystemExit(main(['serve', '--stdio']))",
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    assert process.stdin is not None
    assert process.stdout is not None
    assert process.stderr is not None
    process_stdin = cast(BinaryIO, process.stdin)
    process_stdout = cast(BinaryIO, process.stdout)
    writer = FrameWriter(process_stdin)
    trace: list[dict[str, Any]] = []

    def exchange(message: dict[str, Any]) -> dict[str, Any]:
        trace.append(message)
        writer.write(message)
        response = read_frame(process_stdout)
        assert response is not None
        trace.append(response)
        return response

    initialize = exchange(
        _request(
            1,
            "initialize",
            {
                "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
                "client": {"name": "docwen-e2e", "version": "1.0.0"},
                "features": {"progress": True, "cancellation": True},
            },
        )
    )
    assert initialize["result"]["max_concurrent_tasks"] == 1
    discovery = exchange(_request(2, "capability/list", {}))
    capabilities = {item["capability_id"]: item for item in discovery["result"]["capabilities"]}
    assert set(capabilities) == {
        "convert.markdown.to_docx",
        "convert.markdown.to_xlsx",
        "convert.docx.to_markdown",
        "convert.pdf.to_markdown",
        "convert.ofd.to_markdown",
        "convert.xps.to_markdown",
        "convert.tiff.to_markdown",
        "convert.xlsx.to_csv",
        "convert.xlsx.to_markdown",
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
    assert capabilities["render.pdf.to_png"]["operation"] == "render"
    assert capabilities["convert.png.to_ocr_markdown"]["availability"] == "available"
    assert capabilities["convert.png.to_ocr_markdown"]["dependencies"] == [
        {"dependency_id": "python.pillow", "required": True, "available": True},
        {"dependency_id": "python.rapidocr", "required": True, "available": True},
    ]
    expected_semantic_limitations = list(MACHINE_DOCUMENT_SEMANTICS_LIMITATIONS)
    expected_resolved_document_limitations = list(MACHINE_RESOLVED_DOCUMENT_LIMITATIONS)
    assert capabilities["convert.markdown.to_docx"]["limitations"] == expected_resolved_document_limitations
    assert capabilities["convert.docx.to_markdown"]["limitations"] == expected_semantic_limitations
    assert capabilities["convert.markdown.to_docx"]["input_shape"] == {
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
    assert capabilities["convert.markdown.to_docx"]["output_shape"] == {
        "cardinality": "many",
        "artifact_kinds": ["document", "resource"],
        "relation_types": ["resource_of"],
        "atomic_bundle": True,
    }
    for capability_id in (
        "convert.pdf.to_markdown",
        "convert.ofd.to_markdown",
        "convert.xps.to_markdown",
        "convert.tiff.to_markdown",
    ):
        assert capabilities[capability_id]["output_shape"] == {
            "cardinality": "many",
            "artifact_kinds": ["document", "fragment", "resource"],
            "relation_types": ["fragment_of", "resource_of"],
            "atomic_bundle": True,
            "relation_payloads": ["page_fragment", "page_resource"],
        }
        assert "page_nodes" not in repr(capabilities[capability_id])
    planned = exchange(
        _request(
            3,
            "task/plan",
            {
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
        )
    )
    assert planned["result"]["limitations"] == expected_resolved_document_limitations
    assert list(staging.iterdir()) == []
    accepted = exchange(_request(4, "task/execute", {"plan_id": planned["result"]["plan_id"]}))
    assert accepted["result"]["state"] == "accepted"

    while True:
        notification = read_frame(process_stdout)
        assert notification is not None
        trace.append(notification)
        if notification.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            terminal = notification
            break

    assert terminal["method"] == "task/completed", terminal
    bundle = terminal["params"]["bundle"]
    assert len(bundle["artifacts"]) == 2
    artifact = next(item for item in bundle["artifacts"] if item["kind"] == "document")
    sidecar_artifact = next(item for item in bundle["artifacts"] if item["kind"] == "resource")
    assert sidecar_artifact["media_type"] == ROUND_TRIP_SIDECAR_MEDIA_TYPE
    assert bundle["entries"] == [
        {
            "artifact_id": artifact["artifact_id"],
            "role": "primary",
            "ordinal": 0,
            "preferred": True,
        }
    ]
    assert bundle["relations"] == [
        {
            "type": "resource_of",
            "source_artifact_id": sidecar_artifact["artifact_id"],
            "target_artifact_id": artifact["artifact_id"],
            "role": "manifest",
            "ordinal": 0,
        }
    ]
    output = staging / Path(artifact["locator"])
    sidecar_output = staging / Path(sidecar_artifact["locator"])
    assert output.is_file()
    assert sidecar_output == Path(f"{output}.docwen")
    assert sidecar_artifact["suggested_name"] == f"{artifact['suggested_name']}.docwen"
    assert output.stat().st_size == artifact["size_bytes"]
    assert hashlib.sha256(output.read_bytes()).hexdigest() == artifact["sha256"]
    assert sidecar_output.stat().st_size == sidecar_artifact["size_bytes"]
    assert hashlib.sha256(sidecar_output.read_bytes()).hexdigest() == sidecar_artifact["sha256"]
    sidecar = read_round_trip_sidecar(sidecar_output, docx_path=output)
    assert sidecar.neutral_document == neutral_bytes
    assert sidecar.numbering_export_plan == numbering_plan_bytes
    assert sidecar.authored_source == MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT["document"]["authored_markdown"].encode()
    with zipfile.ZipFile(output) as archive:
        assert "word/document.xml" in archive.namelist()
        assert any(name.startswith("word/media/") for name in archive.namelist())
    verify_machine_document_semantics_docx(output)

    markdown_staging = tmp_path / "markdown-staging"
    markdown_staging.mkdir()
    output_bytes = output.read_bytes()
    markdown_plan = exchange(
        _request(
            5,
            "task/plan",
            {
                "capability_id": "convert.docx.to_markdown",
                "inputs": [
                    {
                        "input_id": "input.2",
                        "kind": "document",
                        "role": "source",
                        "logical_path": "converted/output.docx",
                        "locator": {"kind": "local_path", "path": str(output)},
                        "media_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "size_bytes": len(output_bytes),
                        "sha256": hashlib.sha256(output_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(markdown_staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        )
    )
    assert markdown_plan["result"]["limitations"] == expected_semantic_limitations
    assert markdown_plan["result"]["effective_options"] == {
        "add_numbering": False,
        "image_mode": "file",
        "image_link_style": "wiki_embed",
        "numbering_scheme": "gongwen_standard",
        "ocr_language": "auto",
        "ocr_placement": "main_md",
        "preserve_resources": True,
        "recognize_text": False,
        "remove_numbering": True,
        "table_merge_strategy": "fill",
    }
    assert markdown_plan["result"]["output_shape"] == {
        "cardinality": "many",
        "artifact_kinds": ["document", "fragment", "resource"],
        "relation_types": ["fragment_of", "resource_of"],
        "atomic_bundle": True,
    }
    accepted = exchange(_request(6, "task/execute", {"plan_id": markdown_plan["result"]["plan_id"]}))
    assert accepted["result"]["state"] == "accepted"

    while True:
        notification = read_frame(process_stdout)
        assert notification is not None
        trace.append(notification)
        if notification.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
            markdown_terminal = notification
            break

    assert markdown_terminal["method"] == "task/completed", json.dumps(
        markdown_terminal,
        ensure_ascii=False,
        indent=2,
    )
    markdown_bundle = markdown_terminal["params"]["bundle"]
    markdown_artifact = next(artifact for artifact in markdown_bundle["artifacts"] if artifact["kind"] == "document")
    resource_artifacts = [artifact for artifact in markdown_bundle["artifacts"] if artifact["kind"] == "resource"]
    assert len(resource_artifacts) == 2
    image_artifact = next(artifact for artifact in resource_artifacts if artifact["media_type"] == "image/png")
    manifest_artifact = next(
        artifact
        for artifact in resource_artifacts
        if artifact["media_type"] == "application/vnd.docwen.document-node+json"
    )
    markdown_output = markdown_staging / Path(markdown_artifact["locator"])
    assert markdown_artifact["media_type"] == "text/markdown"
    assert markdown_output.stat().st_size == markdown_artifact["size_bytes"]
    assert hashlib.sha256(markdown_output.read_bytes()).hexdigest() == markdown_artifact["sha256"]
    markdown_text = markdown_output.read_text(encoding="utf-8")
    verify_machine_document_semantics_markdown(markdown_text)
    verify_machine_note_domains_markdown(markdown_text)
    extracted_image = markdown_staging / Path(image_artifact["locator"])
    assert extracted_image.stat().st_size == image_artifact["size_bytes"]
    assert hashlib.sha256(extracted_image.read_bytes()).hexdigest() == image_artifact["sha256"]
    assert markdown_bundle["relations"] == [
        {
            "type": "resource_of",
            "source_artifact_id": image_artifact["artifact_id"],
            "target_artifact_id": markdown_artifact["artifact_id"],
            "role": "image",
            "ordinal": 0,
        },
        {
            "type": "resource_of",
            "source_artifact_id": manifest_artifact["artifact_id"],
            "target_artifact_id": markdown_artifact["artifact_id"],
            "role": "manifest",
            "ordinal": 1,
        },
    ]

    _exercise_auxiliary_capability_matrix(
        tmp_path=tmp_path,
        workbook_source=workbook_source,
        pdf_source=pdf_source,
        tiff_source=tiff_source,
        ocr_source=ocr_source,
        tables_source=tables_source,
        source=source,
        source_bytes=source_bytes,
        process_stdout=process_stdout,
        exchange=exchange,
        trace=trace,
    )

    process_stdin.close()
    assert process.wait(timeout=30) == 0
    assert process.stderr.read() == b""

    validator = MachineContractValidator()
    for message in trace:
        validator.validate_message(message)
    validate_trace(trace, requires_terminal=True)
