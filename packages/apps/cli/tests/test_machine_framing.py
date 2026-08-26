"""Strict Content-Length framing contracts for Machine Protocol v1."""

from __future__ import annotations

import io
import json

import pytest

from docwen_cli.machine.framing import (
    MAX_HEADER_BYTES,
    MAX_MESSAGE_BYTES,
    FrameWriter,
    FramingError,
    read_frame,
)

pytestmark = pytest.mark.contract


def test_frame_writer_and_reader_use_utf8_byte_length() -> None:
    stream = io.BytesIO()
    payload = {"jsonrpc": "2.0", "id": "中文", "method": "capability/list", "params": {}}

    FrameWriter(stream).write(payload)
    raw = stream.getvalue()
    header, body = raw.split(b"\r\n\r\n", 1)

    assert header == f"Content-Length: {len(body)}".encode("ascii")
    assert len(body) != len(body.decode("utf-8"))
    assert read_frame(io.BytesIO(raw)) == payload


@pytest.mark.parametrize(
    ("raw", "code"),
    [
        (b"Content-Length: 2\n\n{}", "frame_header_truncated"),
        (b"content-length: 2\r\n\r\n{}", "invalid_frame_header"),
        (b"Content-Length: 3\r\n\r\n{}", "frame_length_mismatch"),
        (b"Content-Length: 2\r\nX-Test: 1\r\n\r\n{}", "invalid_frame_header"),
    ],
)
def test_malformed_frames_are_rejected(raw: bytes, code: str) -> None:
    with pytest.raises(FramingError) as rejected:
        read_frame(io.BytesIO(raw))
    assert rejected.value.code == code


def test_reader_returns_none_only_for_clean_eof() -> None:
    assert read_frame(io.BytesIO()) is None
    payload = json.dumps(["not", "an", "object"]).encode("utf-8")
    raw = f"Content-Length: {len(payload)}\r\n\r\n".encode("ascii") + payload
    with pytest.raises(FramingError) as rejected:
        read_frame(io.BytesIO(raw))
    assert rejected.value.code == "invalid_frame_payload"


def test_reader_distinguishes_truncated_header_from_header_too_large() -> None:
    with pytest.raises(FramingError) as truncated:
        read_frame(io.BytesIO(b"x" * MAX_HEADER_BYTES))
    assert truncated.value.code == "frame_header_truncated"

    with pytest.raises(FramingError) as oversized:
        read_frame(io.BytesIO(b"x" * (MAX_HEADER_BYTES + 1)))
    assert oversized.value.code == "frame_header_too_large"


def test_reader_rejects_oversized_declared_length_before_reading_body() -> None:
    raw = f"Content-Length: {MAX_MESSAGE_BYTES + 1}\r\n\r\n".encode("ascii")

    with pytest.raises(FramingError) as rejected:
        read_frame(io.BytesIO(raw))

    assert rejected.value.code == "frame_too_large"


def test_writer_rejects_payload_larger_than_machine_message_limit() -> None:
    stream = io.BytesIO()
    payload = {"data": "x" * MAX_MESSAGE_BYTES}

    with pytest.raises(FramingError) as rejected:
        FrameWriter(stream).write(payload)

    assert rejected.value.code == "frame_too_large"
    assert stream.getvalue() == b""


def test_writer_and_reader_accept_exact_machine_message_limit() -> None:
    """Round-trip one 16 MiB JSON body; BytesIO retains one additional framed copy."""

    payload = {"data": ""}
    body_prefix_length = len(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
    payload["data"] = "x" * (MAX_MESSAGE_BYTES - body_prefix_length)
    assert len(json.dumps(payload, separators=(",", ":")).encode("utf-8")) == MAX_MESSAGE_BYTES

    stream = io.BytesIO()
    FrameWriter(stream).write(payload)

    assert read_frame(io.BytesIO(stream.getvalue())) == payload
