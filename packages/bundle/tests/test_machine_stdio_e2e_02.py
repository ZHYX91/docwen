"""Focused tests split from test_machine_stdio_e2e.py."""

from __future__ import annotations

from ._machine_stdio_e2e_support import (
    Any,
    BinaryIO,
    FrameWriter,
    Image,
    MachineContractValidator,
    Path,
    Workbook,
    _create_directory_link,
    _directory_inventory,
    _file_inventory,
    _remove_directory_link,
    _request,
    cast,
    hashlib,
    json,
    load_workbook,
    os,
    pytest,
    read_frame,
    subprocess,
    sys,
    validate_trace,
)


@pytest.mark.e2e
def test_real_stdio_process_executes_new_single_input_capabilities(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# 1. Alpha\n\n| Name | Value |\n|---|---|\n| alpha | 1 |\n", encoding="utf-8")
    proofread_source = tmp_path / "proofread.md"
    proofread_source.write_text("# Proofread\n\nFullwidth digit: １\n", encoding="utf-8")
    import fitz

    pdf_inputs: list[Path] = []
    for index in range(2):
        pdf_path = tmp_path / f"merge-{index + 1}.pdf"
        document = fitz.open()
        document.new_page().insert_text((72, 72), f"Merge page {index + 1}")
        document.save(pdf_path)
        document.close()
        pdf_inputs.append(pdf_path)
    table_input = tmp_path / "table-input.xlsx"
    table_workbook = Workbook()
    table_sheet = table_workbook.active
    assert table_sheet is not None
    table_sheet.append(["beta", 2])
    table_workbook.save(table_input)
    table_workbook.close()
    image_inputs = [tmp_path / "red.png", tmp_path / "blue.png"]
    for image_path, color in zip(image_inputs, ((255, 0, 0), (0, 0, 255)), strict=True):
        with Image.new("RGB", (12, 12), color) as image:
            image.save(image_path)
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
    next_request_id = 1

    def exchange(method: str, params: dict[str, Any]) -> dict[str, Any]:
        nonlocal next_request_id
        request = _request(next_request_id, method, params)
        next_request_id += 1
        trace.append(request)
        writer.write(request)
        response = read_frame(process_stdout)
        assert response is not None
        trace.append(response)
        return response

    def execute(
        capability_id: str,
        input_path: Path,
        media_type: str,
        *,
        options: dict[str, Any] | None = None,
    ) -> tuple[dict[str, Any], Path, dict[str, Any]]:
        return execute_many(capability_id, [(input_path, media_type)], options=options)

    def execute_many(
        capability_id: str,
        inputs: list[tuple[Path, str]],
        *,
        options: dict[str, Any] | None = None,
    ) -> tuple[dict[str, Any], Path, dict[str, Any]]:
        staging = tmp_path / f"staging-{next_request_id}"
        staging.mkdir()
        input_payloads = [(input_path, media_type, input_path.read_bytes()) for input_path, media_type in inputs]
        planned = exchange(
            "task/plan",
            {
                "capability_id": capability_id,
                "inputs": [
                    {
                        "input_id": f"input.{next_request_id}.{index}",
                        "kind": (
                            "document"
                            if media_type
                            in {
                                "text/markdown",
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            }
                            else "resource"
                        ),
                        "role": "source",
                        "logical_path": f"inputs/{index}/{input_path.name}",
                        "locator": {"kind": "local_path", "path": str(input_path)},
                        "media_type": media_type,
                        "size_bytes": len(payload),
                        "sha256": hashlib.sha256(payload).hexdigest(),
                    }
                    for index, (input_path, media_type, payload) in enumerate(input_payloads, start=1)
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(staging)},
                    "staging_policy": "require_empty",
                },
                "options": dict(options or {}),
            },
        )
        accepted = exchange("task/execute", {"plan_id": planned["result"]["plan_id"]})
        assert accepted["result"]["state"] == "accepted"
        while True:
            terminal = read_frame(process_stdout)
            assert terminal is not None
            trace.append(terminal)
            if terminal.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
                break
        assert terminal["method"] == "task/completed", terminal
        bundle = terminal["params"]["bundle"]
        for artifact in bundle["artifacts"]:
            artifact_path = staging / Path(artifact["locator"])
            assert artifact_path.stat().st_size == artifact["size_bytes"]
            assert hashlib.sha256(artifact_path.read_bytes()).hexdigest() == artifact["sha256"]
        return bundle, staging, planned["result"]

    initialized = exchange(
        "initialize",
        {
            "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
            "client": {"name": "docwen-d3-e2e", "version": "1.0.0"},
            "features": {"progress": True, "cancellation": True},
        },
    )
    assert initialized["result"]["protocol"] == {"name": "docwen.machine", "major": 1, "minor": 0}

    xlsx_bundle, xlsx_staging, xlsx_plan = execute(
        "convert.markdown.to_xlsx",
        source,
        "text/markdown",
    )
    assert xlsx_plan["effective_options"] == {}
    assert xlsx_bundle["artifacts"][0]["kind"] == "document"
    xlsx_path = xlsx_staging / Path(xlsx_bundle["artifacts"][0]["locator"])
    workbook = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        assert workbook.sheetnames
    finally:
        workbook.close()

    markdown_bundle, markdown_staging, _ = execute(
        "convert.xlsx.to_markdown",
        xlsx_path,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    markdown_artifact = next(item for item in markdown_bundle["artifacts"] if item["kind"] == "document")
    markdown_path = markdown_staging / Path(markdown_artifact["locator"])
    assert "alpha" in markdown_path.read_text(encoding="utf-8")

    numbering_bundle, numbering_staging, numbering_plan = execute(
        "transform.markdown.heading_numbering",
        source,
        "text/markdown",
        options={
            "remove_numbering": True,
            "add_numbering": True,
            "numbering_scheme": "hierarchical_standard",
        },
    )
    assert numbering_plan["effective_options"]["add_numbering"] is True
    numbered_artifact = numbering_bundle["artifacts"][0]
    assert numbered_artifact["kind"] == "document"
    numbered_path = numbering_staging / Path(numbered_artifact["locator"])
    assert "Alpha" in numbered_path.read_text(encoding="utf-8")

    report_bundle, report_staging, _ = execute(
        "validate.markdown",
        proofread_source,
        "text/markdown",
        options={"enable_sensitive_word": False},
    )
    assert report_bundle["artifacts"][0]["kind"] == "resource"
    assert report_bundle["entries"] == [
        {
            "artifact_id": report_bundle["artifacts"][0]["artifact_id"],
            "role": "supplementary",
            "ordinal": 0,
            "preferred": True,
        }
    ]
    report_path = report_staging / Path(report_bundle["artifacts"][0]["locator"])
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["schema"] == "docwen.proofread_report.v2"
    assert any(issue["matched_text"] == "１" for issue in report["issues"])

    merged_pdf_bundle, merged_pdf_staging, _ = execute_many(
        "merge.pdf.documents",
        [(path, "application/pdf") for path in pdf_inputs],
    )
    assert merged_pdf_bundle["artifacts"][0]["kind"] == "document"
    merged_pdf_path = merged_pdf_staging / Path(merged_pdf_bundle["artifacts"][0]["locator"])
    with fitz.open(merged_pdf_path) as merged_pdf:
        assert merged_pdf.page_count == 2

    partition_bundle, _, partition_plan = execute(
        "split.pdf.partition",
        merged_pdf_path,
        "application/pdf",
        options={"pages": [1]},
    )
    assert partition_plan["effective_options"] == {"split_mode": "custom", "pages": [1]}
    assert [artifact["kind"] for artifact in partition_bundle["artifacts"]] == ["document", "document"]
    assert [(entry["role"], entry["ordinal"]) for entry in partition_bundle["entries"]] == [
        ("section", 0),
        ("section", 1),
    ]

    merged_table_bundle, merged_table_staging, merged_table_plan = execute_many(
        "merge.xlsx.tables",
        [
            (xlsx_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            (table_input, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ],
        options={"merge_mode": "row"},
    )
    assert merged_table_plan["effective_options"]["merge_mode"] == "row"
    merged_table_path = merged_table_staging / Path(merged_table_bundle["artifacts"][0]["locator"])
    merged_workbook = load_workbook(merged_table_path, read_only=True, data_only=True)
    try:
        assert merged_workbook.sheetnames
    finally:
        merged_workbook.close()

    merged_image_bundle, merged_image_staging, _ = execute_many(
        "merge.images.to_tiff",
        [(path, "image/png") for path in image_inputs],
        options={"keep_alpha": False},
    )
    assert merged_image_bundle["artifacts"][0]["kind"] == "resource"
    assert merged_image_bundle["entries"][0]["role"] == "image"
    merged_image_path = merged_image_staging / Path(merged_image_bundle["artifacts"][0]["locator"])
    with Image.open(merged_image_path) as merged_image:
        assert getattr(merged_image, "n_frames", 1) == 2

    process_stdin.close()
    assert process.wait(timeout=30) == 0
    assert process.stderr.read() == b""
    validator = MachineContractValidator()
    for message in trace:
        validator.validate_message(message)
    validate_trace(trace, requires_terminal=True)


@pytest.mark.e2e
def test_real_stdio_process_rejects_linked_staging_then_completes_in_same_session(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# Linked staging safety\n\nFullwidth digit: １\n", encoding="utf-8")
    source_bytes = source.read_bytes()
    source_inventory = _file_inventory(source, relative_to=tmp_path)

    linked_target = tmp_path / "linked-target"
    linked_target.mkdir()
    (linked_target / "sentinel.bin").write_bytes(b"DOCWEN-STAGING-LINK-SENTINEL\x00\xff")
    nested = linked_target / "nested"
    nested.mkdir()
    (nested / "keep.txt").write_text("keep this target unchanged\n", encoding="utf-8")
    target_inventory = _directory_inventory(linked_target)
    assert [item["path"] for item in target_inventory] == ["nested/keep.txt", "sentinel.bin"]

    linked_staging = tmp_path / "linked-staging"
    process: subprocess.Popen[bytes] | None = None
    process_stdin_closed = False
    trace: list[dict[str, Any]] = []
    _create_directory_link(linked_staging, linked_target)
    try:
        if os.name == "nt":
            assert linked_staging.is_junction()
        else:
            assert linked_staging.is_symlink()

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
        next_request_id = 1

        def exchange(method: str, params: dict[str, Any]) -> dict[str, Any]:
            nonlocal next_request_id
            request = _request(next_request_id, method, params)
            next_request_id += 1
            trace.append(request)
            writer.write(request)
            response = read_frame(process_stdout)
            assert response is not None
            trace.append(response)
            return response

        def plan_params(staging_root: Path) -> dict[str, Any]:
            return {
                "capability_id": "validate.markdown",
                "inputs": [
                    {
                        "input_id": "input.link-safety",
                        "kind": "document",
                        "role": "source",
                        "logical_path": "source.md",
                        "locator": {"kind": "local_path", "path": str(source)},
                        "media_type": "text/markdown",
                        "size_bytes": len(source_bytes),
                        "sha256": hashlib.sha256(source_bytes).hexdigest(),
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(staging_root)},
                    "staging_policy": "require_empty",
                },
                "options": {"enable_sensitive_word": False},
            }

        initialized = exchange(
            "initialize",
            {
                "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
                "client": {"name": "docwen-d5-link-e2e", "version": "1.0.0"},
                "features": {"progress": True, "cancellation": True},
            },
        )
        assert initialized["result"]["protocol"] == {"name": "docwen.machine", "major": 1, "minor": 0}

        rejected = exchange("task/plan", plan_params(linked_staging))
        assert "result" not in rejected
        assert rejected == {
            "jsonrpc": "2.0",
            "id": 2,
            "error": {
                "code": -32602,
                "message": "Task planning failed",
                "data": {
                    "task_error": {
                        "category": "security",
                        "code": "staging_root_is_link",
                        "message": "staging root must not be a link or junction",
                        "retryable": False,
                    }
                },
            },
        }
        assert _file_inventory(source, relative_to=tmp_path) == source_inventory
        assert _directory_inventory(linked_target) == target_inventory
        if os.name == "nt":
            assert linked_staging.is_junction()
        else:
            assert linked_staging.is_symlink()
        notification_methods = {"task/progress", "task/completed", "task/failed", "task/cancelled"}
        assert all(message.get("method") not in notification_methods for message in trace)
        assert all(
            "plan_id" not in message.get("result", {}) for message in trace if isinstance(message.get("result"), dict)
        )
        assert all(
            "bundle" not in message.get("params", {}) for message in trace if isinstance(message.get("params"), dict)
        )

        staging = tmp_path / "normal-staging"
        staging.mkdir()
        assert _directory_inventory(linked_target) == target_inventory
        planned = exchange("task/plan", plan_params(staging))
        assert "error" not in planned
        assert list(staging.iterdir()) == []
        assert all(message.get("method") not in notification_methods for message in trace)

        accepted = exchange("task/execute", {"plan_id": planned["result"]["plan_id"]})
        assert accepted["result"]["state"] == "accepted"
        task_id = accepted["result"]["task_id"]
        while True:
            notification = read_frame(process_stdout)
            assert notification is not None
            trace.append(notification)
            if notification.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
                terminal = notification
                break

        task_notifications = [message for message in trace if message.get("method") in notification_methods]
        assert task_notifications
        assert {message["params"]["task_id"] for message in task_notifications} == {task_id}
        terminals = [
            message
            for message in task_notifications
            if message["method"] in {"task/completed", "task/failed", "task/cancelled"}
        ]
        assert terminals == [terminal]
        assert terminal["method"] == "task/completed"

        bundle = terminal["params"]["bundle"]
        assert bundle["task_id"] == task_id
        assert len(bundle["artifacts"]) == 1
        artifact = bundle["artifacts"][0]
        locator = Path(artifact["locator"])
        assert not locator.is_absolute()
        assert ".." not in locator.parts
        output = (staging / locator).resolve(strict=True)
        assert output.is_relative_to(staging.resolve(strict=True))
        output_bytes = output.read_bytes()
        assert len(output_bytes) == artifact["size_bytes"]
        assert hashlib.sha256(output_bytes).hexdigest() == artifact["sha256"]
        report = json.loads(output_bytes)
        assert report["schema"] == "docwen.proofread_report.v2"

        assert _file_inventory(source, relative_to=tmp_path) == source_inventory
        assert _directory_inventory(linked_target) == target_inventory

        process_stdin.close()
        process_stdin_closed = True
        assert process.wait(timeout=30) == 0
        assert process_stdout.read() == b""
        assert process.stderr.read() == b""

        validator = MachineContractValidator()
        for message in trace:
            validator.validate_message(message)
        validate_trace(trace, requires_terminal=True)
    finally:
        if process is not None and process.poll() is None:
            if process.stdin is not None and not process_stdin_closed:
                process.stdin.close()
            try:
                process.wait(timeout=30)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=30)
        if os.path.lexists(linked_staging):
            _remove_directory_link(linked_staging)

    assert not os.path.lexists(linked_staging)
    assert _file_inventory(source, relative_to=tmp_path) == source_inventory
    assert _directory_inventory(linked_target) == target_inventory
