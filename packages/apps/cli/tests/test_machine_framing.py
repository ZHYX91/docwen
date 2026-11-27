"""Strict Content-Length framing contracts for Machine Protocol v2."""

from __future__ import annotations

import io
import json

import pytest

from docwen_cli.machine.framing import (
    MAX_HEADER_BYTES,
    MAX_MESSAGE_BYTES,
    FrameWriter,
    FramingError,
    encode_frame,
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


def test_writer_finishes_short_writes_before_flushing_each_frame() -> None:
    class ShortWriter(io.BytesIO):
        def write(self, data, /):
            return super().write(data[:7])

        def flush(self):
            flushed.append(self.getvalue())

    flushed: list[bytes] = []
    stream = ShortWriter()
    writer = FrameWriter(stream)
    first = {"id": "甲"}
    second = {"id": "乙"}
    writer.write(first)
    writer.write_frame(encode_frame(second))

    assert flushed == [encode_frame(first), encode_frame(first) + encode_frame(second)]
    stream.seek(0)
    assert read_frame(stream) == first
    assert read_frame(stream) == second
    assert read_frame(stream) is None


def test_writer_reports_a_stream_that_stops_accepting_bytes() -> None:
    class StalledWriter(io.BytesIO):
        def write(self, data, /):
            return super().write(data[:1]) if self.tell() == 0 else 0

    with pytest.raises(OSError):
        FrameWriter(StalledWriter()).write({"id": 1})


@pytest.mark.parametrize("literal", ["NaN", "Infinity", "-Infinity"])
def test_non_json_numbers_are_rejected_in_both_directions(literal: str) -> None:
    body = ('{"value":' + literal + "}").encode("ascii")
    raw = f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body
    with pytest.raises(FramingError) as incoming:
        read_frame(io.BytesIO(raw))
    assert incoming.value.code == "invalid_frame_payload"

    stream = io.BytesIO()
    with pytest.raises(FramingError) as outgoing:
        FrameWriter(stream).write({"value": float(literal)})
    assert outgoing.value.code == "invalid_frame_payload"
    assert stream.getvalue() == b""
