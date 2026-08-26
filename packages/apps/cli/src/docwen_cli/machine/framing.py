"""Strict Content-Length framing for DocWen Machine Protocol v1."""

from __future__ import annotations

import json
import re
import threading
from typing import Any, BinaryIO

MAX_MESSAGE_BYTES = 16 * 1024 * 1024
MAX_HEADER_BYTES = 64
_HEADER = re.compile(rb"Content-Length: ([1-9][0-9]*)\r\n\r\n")


class FramingError(ValueError):
    """The input stream did not contain one canonical complete frame."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def read_frame(stream: BinaryIO) -> dict[str, Any] | None:
    """Read one frame; return ``None`` only for clean EOF between frames."""

    header = bytearray()
    while not header.endswith(b"\r\n\r\n"):
        byte = stream.read(1)
        if not byte:
            if not header:
                return None
            raise FramingError("frame_header_truncated", "frame header ended before CRLF CRLF")
        header.extend(byte)
        if len(header) > MAX_HEADER_BYTES:
            raise FramingError("frame_header_too_large", "frame header exceeds the protocol limit")

    match = _HEADER.fullmatch(bytes(header))
    if match is None:
        raise FramingError("invalid_frame_header", "frame must contain one canonical Content-Length header")
    content_length = int(match.group(1))
    if content_length > MAX_MESSAGE_BYTES:
        raise FramingError("frame_too_large", f"frame exceeds {MAX_MESSAGE_BYTES} bytes")
    body = _read_exact(stream, content_length)
    try:
        payload = json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise FramingError("invalid_frame_payload", "frame body must be one UTF-8 JSON object") from exc
    if not isinstance(payload, dict):
        raise FramingError("invalid_frame_payload", "frame body must be one JSON object")
    return payload


def _read_exact(stream: BinaryIO, length: int) -> bytes:
    chunks = bytearray()
    while len(chunks) < length:
        chunk = stream.read(length - len(chunks))
        if not chunk:
            raise FramingError("frame_length_mismatch", "declared Content-Length exceeds available bytes")
        chunks.extend(chunk)
    return bytes(chunks)


class FrameWriter:
    """Serialize complete JSON frames atomically across task worker threads."""

    def __init__(self, stream: BinaryIO) -> None:
        self._stream = stream
        self._lock = threading.Lock()

    def write(self, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        if len(body) > MAX_MESSAGE_BYTES:
            raise FramingError("frame_too_large", f"frame exceeds {MAX_MESSAGE_BYTES} bytes")
        frame = f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body
        with self._lock:
            self._stream.write(frame)
            self._stream.flush()


__all__ = ["MAX_MESSAGE_BYTES", "FrameWriter", "FramingError", "read_frame"]
