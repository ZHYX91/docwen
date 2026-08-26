"""Data URI image detection and temporary file materialisation.

Covers F-H2-006: ``is_data_uri_image`` and ``resolve_data_uri_image_to_temp_file``.

These utilities provide a single shared implementation so that every plugin
does not need to replicate its own data-URI scanning or decoding logic.
"""

from __future__ import annotations

import base64
import contextlib
import logging
import os
import re
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)

_MAX_DATA_URI_IMAGE_BYTES = 10 * 1024 * 1024

_DATA_URI_IMAGE_MIME_TO_EXT: dict[str, str] = {
    "png": ".png",
    "jpeg": ".jpg",
    "jpg": ".jpg",
    "gif": ".gif",
    "bmp": ".bmp",
    "tiff": ".tiff",
    "webp": ".webp",
    "svg+xml": ".svg",
    "icon": ".ico",
}


def is_data_uri_image(link_target: str) -> bool:
    """Return ``True`` when *link_target* is a base64-encoded data URI image.

    Detects the ``data:image/<mime>;base64,...`` pattern commonly found in
    HTML ``<img src="...">`` attributes and inline Markdown image links.
    """
    return link_target.startswith("data:image/") and ";base64," in link_target


def _estimate_base64_decoded_size(payload: str) -> int:
    """Estimate the decoded byte size of a base64 *payload* string."""
    payload_len = len(payload)
    if payload_len == 0:
        return 0
    padding = 2 if payload.endswith("==") else 1 if payload.endswith("=") else 0
    return max(0, (payload_len * 3) // 4 - padding)


def resolve_data_uri_image_to_temp_file(data_uri: str, *, temp_dir: str | None = None) -> str | None:
    """Decode a base64 data URI image and write it to a temporary file.

    The temporary file is created with ``mkstemp`` in *temp_dir* (or the
    system default).  The file suffix is derived from the MIME subtype in the
    data URI header (e.g. ``.png`` for ``image/png``).

    Returns the absolute path of the temp file on success, or ``None`` when
    the data URI is malformed, the payload exceeds the size limit, or
    decoding / writing fails.
    """
    try:
        header, payload = data_uri.split(",", 1)
    except ValueError:
        return None

    if not header.startswith("data:image/") or ";base64" not in header:
        return None

    mime_subtype = header[len("data:image/") :].split(";", 1)[0].strip().lower()
    ext = _DATA_URI_IMAGE_MIME_TO_EXT.get(mime_subtype)
    if not ext:
        cleaned = re.sub(r"[^a-z0-9]+", "", mime_subtype)
        ext = f".{cleaned}" if cleaned else ".img"

    payload = payload.strip()
    payload = payload.translate({ord(c): None for c in " \r\n\t"})

    if _estimate_base64_decoded_size(payload) > _MAX_DATA_URI_IMAGE_BYTES:
        logger.warning(
            "data URI image exceeds size limit (max: %.2f MB)",
            _MAX_DATA_URI_IMAGE_BYTES / (1024 * 1024),
        )
        return None

    try:
        image_bytes = base64.b64decode(payload, validate=True)
    except Exception as e:
        logger.warning("data URI image base64 decode failed: %s", e)
        return None

    if len(image_bytes) > _MAX_DATA_URI_IMAGE_BYTES:
        logger.warning(
            "data URI image exceeds size limit (max: %.2f MB)",
            _MAX_DATA_URI_IMAGE_BYTES / (1024 * 1024),
        )
        return None

    fd = None
    temp_path: str | None = None
    try:
        fd, temp_path = tempfile.mkstemp(suffix=ext, prefix="docwen_data_uri_", dir=temp_dir)
        with os.fdopen(fd, "wb") as f:
            fd = None
            f.write(image_bytes)
        logger.info(
            "data URI image decoded to temp file: %s (%.2f KB)",
            Path(temp_path).name,
            len(image_bytes) / 1024,
        )
        return temp_path
    except Exception as e:
        logger.error("data URI image temp file write failed: %s", e)
        if fd is not None:
            with contextlib.suppress(Exception):
                os.close(fd)
        if temp_path is not None:
            with contextlib.suppress(Exception):
                Path(temp_path).unlink(missing_ok=True)
        return None
