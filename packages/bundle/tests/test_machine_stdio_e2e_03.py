"""Focused tests split from test_machine_stdio_e2e.py."""

from __future__ import annotations

from ._machine_stdio_e2e_support import (
    BinaryIO,
    FrameWriter,
    Path,
    _create_directory_link,
    _directory_inventory,
    _remove_directory_link,
    _request,
    cast,
    hashlib,
    pytest,
    read_frame,
    subprocess,
    sys,
)


@pytest.mark.e2e
def test_real_stdio_process_rejects_linked_input_without_reading_target(tmp_path: Path) -> None:
    target = tmp_path / "linked-input-target"
    target.mkdir()
    source_target = target / "source.md"
    source_target.write_text("# linked input sentinel\n", encoding="utf-8")
    target_inventory = _directory_inventory(target)
    link = tmp_path / "linked-input"
    _create_directory_link(link, target)
    source_link = link / "source.md"
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
    try:
        assert process.stdin is not None
        assert process.stdout is not None
        writer = FrameWriter(cast(BinaryIO, process.stdin))
        source_bytes = source_target.read_bytes()
        writer.write(
            _request(
                1,
                "initialize",
                {
                    "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
                    "client": {"name": "docwen-linked-input-e2e", "version": "1.0.0"},
                    "features": {"progress": True, "cancellation": True},
                },
            )
        )
        initialized = read_frame(cast(BinaryIO, process.stdout))
        assert initialized is not None
        assert initialized["result"]["protocol"] == {"name": "docwen.machine", "major": 1, "minor": 0}
        writer.write(
            _request(
                2,
                "task/plan",
                {
                    "capability_id": "validate.markdown",
                    "inputs": [
                        {
                            "input_id": "input.link",
                            "kind": "document",
                            "role": "source",
                            "logical_path": "source.md",
                            "locator": {"kind": "local_path", "path": str(source_link)},
                            "media_type": "text/markdown",
                            "size_bytes": len(source_bytes),
                            "sha256": hashlib.sha256(source_bytes).hexdigest(),
                        }
                    ],
                    "output": {
                        "staging_root": {"kind": "local_path", "path": str(staging)},
                        "staging_policy": "require_empty",
                    },
                    "options": {"enable_sensitive_word": False},
                },
            )
        )
        rejected = read_frame(cast(BinaryIO, process.stdout))
        assert rejected is not None
        assert rejected["error"]["code"] == -32602
        assert rejected["error"]["data"]["task_error"] == {
            "category": "security",
            "code": "input_is_link",
            "message": "input must not be a link or junction",
            "retryable": False,
        }
        assert _directory_inventory(target) == target_inventory
        assert list(staging.iterdir()) == []
    finally:
        if process.stdin is not None and not process.stdin.closed:
            process.stdin.close()
        if process.poll() is None:
            process.wait(timeout=30)
        assert process.stderr is not None
        assert process.stderr.read() == b""
        _remove_directory_link(link)
