"""Managed OFD/XPS normalization for Application-owned PDF actions.

The PDF merger and splitter continue to receive only validated PDF inputs.
The controller owns the protective snapshot and the temporary output lifetime.
"""

from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_application.preconversion.pre_converter import (
    PreConversionFailure,
    PreConversionResult,
    _is_cancel_requested,
    _prepare_protective_snapshot,
)
from docwen_core.ofd import apply_easyofd_patches, easyofd_import_boundary

if TYPE_CHECKING:
    from docwen_core.protocols import CancellationTokenView


def pre_convert_layout(
    input_path: str,
    source_format: str,
    *,
    staging_dir: str,
    cancel: CancellationTokenView | None = None,
) -> PreConversionResult | PreConversionFailure:
    """Normalize one admitted OFD/XPS source without publishing any output."""
    fmt = source_format.strip().lower()
    if fmt not in {"ofd", "xps"}:
        return PreConversionFailure(
            message="Only admitted OFD/XPS can be normalized for a PDF action.",
            error_type="invalid_input",
            diagnostic_code="LAYOUT-ACTION-SOURCE-UNSUPPORTED",
        )
    if _is_cancel_requested(cancel):
        return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")

    source = Path(input_path)
    work = Path(staging_dir)
    protected = work / f"input.{fmt}"
    destination = work / f"{source.stem}.pdf"
    created_at = datetime.now().astimezone().isoformat()
    try:
        _prepare_protective_snapshot(source, protected, cancel)
        digest = hashlib.sha256()
        with protected.open("rb") as stream:
            while block := stream.read(1024 * 1024):
                if _is_cancel_requested(cancel):
                    return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")
                digest.update(block)
        if fmt == "ofd":
            with easyofd_import_boundary():
                from easyofd import OFD  # type: ignore[import-untyped]

            apply_easyofd_patches()
            ofd = OFD()
            ofd.read(str(protected), fmt="path")
            pdf_data = ofd.to_pdf()
            backend = "easyofd"
        else:
            import fitz

            with fitz.open(str(protected), filetype="xps") as source_doc:
                pdf_data = source_doc.convert_to_pdf()
            backend = "PyMuPDF"
        if _is_cancel_requested(cancel):
            return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")
        import fitz

        if not isinstance(pdf_data, (bytes, bytearray)) or not pdf_data.startswith(b"%PDF-"):
            raise ValueError("Layout normalization produced non-PDF content")
        with fitz.open(stream=pdf_data, filetype="pdf") as normalized:
            if normalized.page_count <= 0:
                raise ValueError("Layout normalization produced an empty PDF")
        destination.write_bytes(pdf_data)
        if _is_cancel_requested(cancel):
            return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")
    except ImportError as exc:
        return PreConversionFailure(
            message=f"Required layout PDF dependency is unavailable: {exc}",
            error_type="dependency_missing",
            diagnostic_code="LAYOUT-ACTION-DEPENDENCY-MISSING",
        )
    except Exception as exc:
        if _is_cancel_requested(cancel):
            return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")
        return PreConversionFailure(
            message=f"Cannot normalize {fmt.upper()} for PDF operation: {exc}",
            error_type="conversion_failed",
            diagnostic_code="LAYOUT-ACTION-PRECONVERSION-FAILED",
        )

    return PreConversionResult(
        pre_converted_path=str(destination),
        original_source_format=fmt,
        backend=backend,
        source_sha256=digest.hexdigest(),
        created_at=created_at,
    )
