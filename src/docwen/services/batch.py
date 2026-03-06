from __future__ import annotations

import concurrent.futures
import threading
import time
from collections.abc import Callable
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

from docwen.formats import category_from_actual_format
from docwen.utils.file_type_utils import detect_actual_file_format

from .cancellation import CancellationToken
from .error_codes import ERROR_CODE_SKIPPED_SAME_FORMAT, ERROR_CODE_UNKNOWN_ERROR
from .requests import BatchRequest, ConversionRequest
from .result import ConversionResult


@dataclass(slots=True)
class BatchEvent:
    type: str
    file_path: str
    index: int
    total: int
    result: FileResult | None = None


@dataclass(slots=True)
class FileResult:
    file_path: str
    success: bool
    status: str
    output_path: str | None = None
    message: str | None = None
    error_code: str | None = None
    details: str | None = None
    duration_s: float | None = None


@dataclass(slots=True)
class BatchResult:
    total: int
    processed_count: int
    results: list[FileResult]
    success_count: int
    failed_count: int
    cancelled: bool


def execute_batch(
    batch_request: BatchRequest,
    *,
    execute_one: Callable[[ConversionRequest], ConversionResult],
    cancel_token: CancellationToken | None = None,
    event_sink: Callable[[BatchEvent], None] | None = None,
) -> BatchResult:
    cancel_token = cancel_token or CancellationToken()
    total = len(batch_request.requests)

    office_lock = threading.Lock()
    layout_preprocess_lock = threading.Lock()
    ocr_semaphore = threading.Semaphore(1)

    def _emit(event: BatchEvent) -> None:
        if event_sink:
            event_sink(event)

    @contextmanager
    def _with_semaphore(sem: threading.Semaphore):
        sem.acquire()
        try:
            yield
        finally:
            sem.release()

    def _worker(index: int, req: ConversionRequest) -> FileResult:
        if cancel_token.is_cancelled:
            message = "操作已取消"
            translator = req.options.get("t")
            if callable(translator):
                message = str(translator("conversion.messages.operation_cancelled", default=message))
            return FileResult(
                file_path=req.file_path or "",
                success=False,
                status="cancelled",
                message=message,
            )

        file_path = req.file_path or ""
        _emit(BatchEvent(type="file_started", file_path=file_path, index=index, total=total))

        def _run() -> ConversionResult:
            return execute_one(req)

        if not req.actual_format and req.file_path:
            try:
                req.actual_format = detect_actual_file_format(req.file_path)
            except Exception:
                ext = Path(req.file_path).suffix.lower().lstrip(".")
                if ext == "markdown":
                    ext = "md"
                req.actual_format = ext or "unknown"
        if not req.category and req.file_path:
            req.category = category_from_actual_format(req.actual_format)

        actual_format = (req.actual_format or "").lower()
        category = (req.category or "").lower()
        needs_office_lock = (category == "document" and actual_format in {"doc", "wps", "rtf", "odt"}) or (
            category == "spreadsheet" and actual_format in {"xls", "et", "ods"}
        )
        needs_layout_lock = category == "layout" and actual_format in {"ofd", "xps", "caj"}
        needs_ocr_lock = bool(req.options.get("extract_ocr"))

        start = time.perf_counter()
        try:
            if needs_office_lock:
                with office_lock:
                    result = _run()
            elif needs_layout_lock:
                with layout_preprocess_lock:
                    result = _run()
            elif needs_ocr_lock:
                with _with_semaphore(ocr_semaphore):
                    result = _run()
            else:
                result = _run()
        except Exception as e:
            result = ConversionResult.fail(message=str(e), error=e, error_code=ERROR_CODE_UNKNOWN_ERROR, details=str(e))
        duration_s = max(0.0, time.perf_counter() - start)

        if cancel_token.is_cancelled:
            status = "cancelled"
        elif result.success and result.error_code == ERROR_CODE_SKIPPED_SAME_FORMAT:
            status = "skipped"
        else:
            status = "completed" if result.success else "failed"

        file_result = FileResult(
            file_path=file_path,
            success=result.success,
            status=status,
            output_path=result.output_path,
            message=result.message,
            error_code=result.error_code,
            details=result.details,
            duration_s=duration_s,
        )
        _emit(BatchEvent(type="file_finished", file_path=file_path, index=index, total=total, result=file_result))
        return file_result

    results_by_index: list[FileResult | None] = [None] * total
    completed_count = 0

    max_workers = max(1, int(batch_request.max_workers or 1))
    if not batch_request.continue_on_error:
        max_workers = 1

    if max_workers == 1:
        for index, req in enumerate(batch_request.requests, start=1):
            if cancel_token.is_cancelled:
                break
            completed_count = index
            fr = _worker(index, req)
            results_by_index[index - 1] = fr
            if (not fr.success) and (not batch_request.continue_on_error):
                break
    else:
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(_worker, index, req): (index, req)
                for index, req in enumerate(batch_request.requests, start=1)
            }
            try:
                for fut in concurrent.futures.as_completed(futures):
                    if cancel_token.is_cancelled:
                        break
                    index, _req = futures[fut]
                    fr = fut.result()
                    results_by_index[index - 1] = fr
                    completed_count += 1
            finally:
                if cancel_token.is_cancelled:
                    try:
                        executor.shutdown(wait=False, cancel_futures=True)
                    except TypeError:
                        executor.shutdown(wait=False)

    if cancel_token.is_cancelled:
        for i, req in enumerate(batch_request.requests, start=1):
            if results_by_index[i - 1] is not None:
                continue
            file_path = req.file_path or ""
            message = "操作已取消"
            translator = req.options.get("t")
            if callable(translator):
                message = str(translator("conversion.messages.operation_cancelled", default=message))
            results_by_index[i - 1] = FileResult(
                file_path=file_path,
                success=False,
                status="cancelled",
                message=message,
            )

    results = [r for r in results_by_index if r is not None]
    success_count = sum(1 for r in results if r.success)
    failed_count = sum(1 for r in results if not r.success and r.status != "cancelled")
    return BatchResult(
        total=total,
        processed_count=min(total, max(0, int(completed_count))),
        results=results,
        success_count=success_count,
        failed_count=failed_count,
        cancelled=cancel_token.is_cancelled,
    )
