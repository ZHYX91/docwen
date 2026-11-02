"""Independent oracles and transport ownership for the optional packaged gate."""

from __future__ import annotations

import io
import json
import os
import queue
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image
from scripts.release import conversion_fidelity_smoke as smoke

pytestmark = pytest.mark.unit


def test_frame_reader_handles_short_reads_and_rejects_oversized_frames():
    class Fragmented(io.BytesIO):
        def read(self, size=-1):
            return super().read(min(size, 3) if size >= 0 else 3)

    value = {"jsonrpc": "2.0", "result": {"text": "中文😀"}}
    body = json.dumps(value, ensure_ascii=False).encode()
    stream = Fragmented(f"Content-Length: {len(body)}\r\n\r\n".encode() + body)
    assert smoke.read_frame(stream) == value
    assert smoke.read_frame(stream) is None
    with pytest.raises(RuntimeError, match="exceeds_wire_limit"):
        smoke.read_frame(io.BytesIO(f"Content-Length: {smoke.MAX_FRAME_BYTES + 1}\r\n\r\n".encode()))
    with pytest.raises(EOFError):
        smoke.read_frame(io.BytesIO(b"Content-Length: 8\r\n\r\n{}"))


def test_native_entry_rejects_source_launcher_and_isolates_config(tmp_path, monkeypatch):
    binary = tmp_path / ("DocWenCLI.exe" if os.name == "nt" else "DocWenCLI")
    binary.write_text("#!/usr/bin/python\nimport docwen_cli\n")
    with pytest.raises(RuntimeError, match="native_packaged_cli"):
        smoke.packaged_binary(tmp_path)
    monkeypatch.setenv("PYTHONPATH", "old-editable-source")
    monkeypatch.setenv("PYTHONHOME", "old-python")
    monkeypatch.setenv("DOCWEN_CONFIG_DIR", "user-config")
    env = smoke.isolated_environment(tmp_path)
    assert "PYTHONPATH" not in env and "PYTHONHOME" not in env
    assert env["DOCWEN_CONFIG_DIR"] == str(tmp_path / "config")
    assert (tmp_path / "config" / "image.toml").read_text() == 'compress_mode = "lossless"\n'


def test_failed_process_creation_releases_owned_transcript_files(tmp_path, monkeypatch):
    def fail_start(*_args, **_kwargs):
        raise OSError("controlled launch failure")

    monkeypatch.setattr(smoke.subprocess, "Popen", fail_start)
    with pytest.raises(OSError, match="controlled launch failure"):
        smoke.Machine(Path("missing-binary"), tmp_path, os.environ.copy(), timeout=1)
    # Windows refuses these renames when the failed constructor leaks handles.
    for name in ("machine.stderr.txt", "machine.requests.jsonl", "machine.responses.jsonl"):
        path = tmp_path / name
        path.rename(path.with_suffix(".closed"))


def test_continuously_queued_progress_cannot_extend_request_deadline(monkeypatch):
    machine = smoke.Machine.__new__(smoke.Machine)
    machine.timeout = 1
    machine.request_id = 0
    machine.requests = io.StringIO()
    machine.process = SimpleNamespace(stdin=io.BytesIO())
    machine.messages = []
    machine.pending = queue.Queue()
    for sequence in range(20):
        machine.pending.put({"method": "task/progress", "params": {"sequence": sequence}})
    machine.pending.put({"id": 1, "result": {"too_late": True}})
    clock = iter(index / 4 for index in range(100))
    monkeypatch.setattr(smoke.time, "monotonic", lambda: next(clock))
    with pytest.raises(TimeoutError, match="packaged_machine_response_timeout"):
        machine.request("initialize", {})
    assert len(machine.messages) < 20


@pytest.mark.parametrize("clock_values", [(1.0,), (0.0, 1.0)])
def test_deadline_is_enforced_before_and_after_reading_queued_message(monkeypatch, clock_values):
    machine = smoke.Machine.__new__(smoke.Machine)
    machine.pending = queue.Queue()
    machine.pending.put({"method": "task/completed"})
    machine.messages = []
    clock = iter(clock_values)
    monkeypatch.setattr(smoke.time, "monotonic", lambda: next(clock))
    with pytest.raises(TimeoutError, match="packaged_machine_response_timeout"):
        machine._next(1.0)
    assert not machine.messages


def test_cli_timeout_retains_partial_output_bytes(tmp_path, monkeypatch):
    failure = smoke.subprocess.TimeoutExpired("controlled", 1, output=b'{"partial":', stderr=b"diagnostic\n")

    def timeout(*_args, **_kwargs):
        raise failure

    monkeypatch.setattr(smoke.subprocess, "run", timeout)
    with pytest.raises(smoke.subprocess.TimeoutExpired) as caught:
        smoke.cli_convert(Path("packaged-binary"), tmp_path / "source.png", "webp", tmp_path, {}, 1)
    assert caught.value is failure
    assert (tmp_path / "source.stdout.json").read_bytes() == b'{"partial":'
    assert (tmp_path / "source.stderr.txt").read_bytes() == b"diagnostic\n"


def test_request_timeout_stops_only_owned_process_when_stdin_is_not_read(tmp_path, monkeypatch):
    original = smoke.subprocess.Popen

    def spawn(command, **kwargs):
        assert command == ["packaged-binary", "serve", "--stdio"]
        return original([sys.executable, "-u", "-c", "import time; time.sleep(30)"], **kwargs)

    monkeypatch.setattr(smoke.subprocess, "Popen", spawn)
    machine = smoke.Machine(Path("packaged-binary"), tmp_path, os.environ.copy(), timeout=0.2)
    try:
        with pytest.raises(TimeoutError, match="packaged_machine_request_timeout"):
            machine.request("initialize", {"large": "x" * (1024 * 1024)})
        assert machine.process.poll() is not None
    finally:
        with pytest.raises(RuntimeError, match="packaged_machine_exit_failed"):
            machine.close()
    assert not machine.reader.is_alive()
    assert (tmp_path / "machine.requests.jsonl").stat().st_size > 1024 * 1024


def test_saved_spreadsheet_oracle_checks_values_and_sparse_cell_count(tmp_path):
    path = tmp_path / "sheet.xlsx"
    smoke.write_sheet(path, {"A1": "NaN,1", "SR512": 4})
    smoke.verify_sheet(path, {"A1": "NaN,1", "SR512": 4})
    with pytest.raises(RuntimeError, match="value_mismatch"):
        smoke.verify_sheet(path, {"A1": None, "SR512": 4})
    with pytest.raises(RuntimeError, match="materialized_cells"):
        smoke.verify_sheet(path, {"A1": "NaN,1"})


@pytest.mark.parametrize("mode", ["RGB", "RGBA"])
def test_webp_oracle_rejects_lossy_pixels_including_transparent_rgb(tmp_path, mode):
    source = tmp_path / "source.png"
    output = tmp_path / "result.webp"
    with Image.new(mode, (32, 32)) as image:
        for y in range(32):
            for x in range(32):
                color = ((x * 47 + y * 13) % 256, (x * 17 + y * 59) % 256, (x * 79 + y * 11) % 256)
                image.putpixel((x, y), (*color, x % 2 * 255) if mode == "RGBA" else color)
        image.save(source)
        image.save(output, lossless=True, exact=True)
        smoke.verify_webp(output, source)
        image.save(output, quality=95)
        with pytest.raises(RuntimeError, match="not_lossless"):
            smoke.verify_webp(output, source)


def test_tiff_oracle_rejects_lost_later_frames(tmp_path):
    output = tmp_path / "result.tif"
    frames = [Image.new("RGBA", (2, 2), color) for color in smoke.FRAME_COLORS]
    try:
        frames[0].save(output, save_all=True, append_images=frames[1:])
        smoke.verify_tiff(output, alpha=True)
        frames[0].save(output, save_all=True, append_images=frames[1:2])
        with pytest.raises(RuntimeError, match="missing_frames"):
            smoke.verify_tiff(output, alpha=True)
    finally:
        for frame in frames:
            frame.close()


def test_presentation_oracle_uses_rendered_literal_text_and_shape_occurrence(tmp_path):
    literal = "&lt;b&gt;literal&lt;/b&gt; &#91;manual&#93;(https://example.com) &#42;not italic&#42; &amp;copy;"
    body = "\n\n".join(
        (
            "# Guide &#42;literal&#42;",
            f"## {literal}",
            literal,
            "&#35; literal heading",
            "&#45; not a bullet",
            "1&#46; not ordered",
            "&#45;&#45;&#45;",
            "$$price$$",
        )
    )
    output = tmp_path / "result.md"
    output.write_text(body, encoding="utf-8")
    smoke.verify_presentation(output)
    output.write_text(body + f"\n\n{literal}", encoding="utf-8")
    with pytest.raises(RuntimeError, match="title_duplicated"):
        smoke.verify_presentation(output)
    output.write_text(body.replace(literal, smoke.LITERAL), encoding="utf-8")
    with pytest.raises(RuntimeError, match="title_changed"):
        smoke.verify_presentation(output)


def test_bundle_oracle_checks_hash_and_staging_boundary(tmp_path):
    staging = tmp_path / "staging"
    staging.mkdir()
    output = staging / "result.xlsx"
    output.write_bytes(b"fixture")
    artifact = {**smoke.identity(output), "locator": "result.xlsx", "media_type": smoke.XLSX_MEDIA}
    terminal = {
        "method": "task/completed",
        "params": {
            "task_id": "t",
            "bundle": {
                "schema": "docwen.artifact_bundle.v3",
                "task_id": "t",
                "artifacts": [artifact],
            },
        },
    }
    assert smoke.output_from_terminal(terminal, staging, smoke.XLSX_MEDIA) == output
    artifact["sha256"] = "0" * 64
    with pytest.raises(RuntimeError, match="integrity"):
        smoke.output_from_terminal(terminal, staging, smoke.XLSX_MEDIA)
    outside = tmp_path / "other.xlsx"
    outside.write_bytes(b"fixture")
    artifact.update(smoke.identity(outside), locator="../other.xlsx")
    with pytest.raises(RuntimeError, match="outside_staging"):
        smoke.output_from_terminal(terminal, staging, smoke.XLSX_MEDIA)


@pytest.mark.parametrize("duplicate", [False, True])
def test_owned_transport_observes_one_terminal_through_normal_eof(tmp_path, monkeypatch, duplicate):
    # This controlled subprocess tests the harness, not the DocWen product.
    code = """
import json, sys
def send(value):
    raw=json.dumps(value).encode()
    sys.stdout.buffer.write(f'Content-Length: {len(raw)}\\r\\n\\r\\n'.encode()+raw)
    sys.stdout.buffer.flush()
while line:=sys.stdin.buffer.readline():
    length=int(line.split(b':')[1]); sys.stdin.buffer.readline()
    request=json.loads(sys.stdin.buffer.read(length))
    result={'plan_id':'p'} if request['method']=='task/plan' else {'state':'accepted','task_id':'t'}
    send({'jsonrpc':'2.0','id':request['id'],'result':result})
    if request['method']=='task/execute':
        send({'jsonrpc':'2.0','method':'task/progress','params':{'task_id':'t','sequence':1}})
        send({'jsonrpc':'2.0','method':'task/completed','params':{'task_id':'t','sequence':2}})
        if DUPLICATE:
            send({'jsonrpc':'2.0','method':'task/completed','params':{'task_id':'t','sequence':3}})
""".replace("DUPLICATE", repr(duplicate))
    original = smoke.subprocess.Popen

    def spawn(command, **kwargs):
        assert command == ["packaged-binary", "serve", "--stdio"]
        return original([sys.executable, "-u", "-c", code], **kwargs)

    monkeypatch.setattr(smoke.subprocess, "Popen", spawn)
    source = tmp_path / "a.xlsx"
    source.write_bytes(b"synthetic")
    machine = smoke.Machine(Path("packaged-binary"), tmp_path, os.environ.copy(), timeout=10)
    terminal = machine.task("transport", [source], tmp_path / "out", "fixture", {})
    assert terminal["method"] == "task/completed"
    if duplicate:
        with pytest.raises(RuntimeError, match="terminal_not_unique"):
            machine.close()
    else:
        machine.close()
    assert machine.process.poll() == 0
    assert not machine.reader.is_alive()
