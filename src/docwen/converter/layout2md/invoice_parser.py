import threading
from typing import Any

from .invoice_cn import convert_invoice_cn_layout_to_md


def convert_invoice_layout_to_md(
    *,
    file_path: str,
    actual_format: str,
    output_dir: str,
    basename_for_output: str,
    original_file_stem: str | None = None,
    cancel_event: threading.Event | None = None,
    progress_callback: Any | None = None,
) -> dict[str, Any]:
    return convert_invoice_cn_layout_to_md(
        file_path=file_path,
        actual_format=actual_format,
        output_dir=output_dir,
        basename_for_output=basename_for_output,
        original_file_stem=original_file_stem,
        cancel_event=cancel_event,
        progress_callback=progress_callback,
    )
