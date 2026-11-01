"""Phase instrumentation around a real composed CLI, for subprocess tests only.

The Application, CSV converter, Runtime and finalizer remain production code.
Gates only delay a worker or a concrete before/after-publication boundary.
"""

from __future__ import annotations

import json
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import replace
from pathlib import Path
from typing import Any


def main(mode: str, root: Path) -> int:
    from docwen_application.controller import ApplicationController
    from docwen_bundle.cli_entry import main as cli_main
    from docwen_cli.execution_deadline import ExecutionDeadline
    from docwen_plugin_spreadsheet.csv_xlsx.converter import CsvToXlsxConverter
    from docwen_runtime.output.finalizer import OutputFinalizer

    trace_lock = threading.Lock()
    ready_count = 0
    submitted_count = 0
    trace_path = root / "trace.jsonl"
    requests_by_id = {}
    reservation_windows = {"prepare_return", "deadline_return", "alias_published"}

    def trace(event: str, **details: Any) -> None:
        with trace_lock, trace_path.open("a", encoding="utf-8") as stream:
            stream.write(json.dumps({"event": event, "time": time.monotonic(), **details}) + "\n")

    def wait_release(cancellation: Any = None) -> None:
        end = time.monotonic() + 6
        while not (root / "release").exists() and time.monotonic() < end:
            if cancellation is not None and cancellation.is_cancelled:
                return
            time.sleep(0.01)

    def wait_cancellation(cancellation: Any) -> None:
        end = time.monotonic() + 6
        while not cancellation.is_cancelled and time.monotonic() < end:
            time.sleep(0.01)
        trace("observed-cancellation", cancelled=cancellation.is_cancelled)

    original_prepare = ApplicationController.prepare_execution_cancellation
    original_cancel = ApplicationController.cancel
    original_execute = ApplicationController.execute_single
    original_stop = ApplicationController.stop
    original_register = ExecutionDeadline.register
    original_pool_init = ThreadPoolExecutor.__init__
    original_submit = ThreadPoolExecutor.submit
    original_convert = CsvToXlsxConverter.convert
    original_finalize = OutputFinalizer.finalize

    class AliasGateMap(dict):
        armed = True

        def __setitem__(self, key, value):
            super().__setitem__(key, value)
            if self.armed:
                self.armed = False
                trace("alias-published", task_id=key, alias_count=len(self))
                trace("ready")
                wait_release()

    def prepare(self, request, *, batch=False, on_reserved=None):
        requests_by_id[request.request_id] = request
        if mode == "alias_published" and not isinstance(self._cancellation_scopes, AliasGateMap):
            self._cancellation_scopes = AliasGateMap(self._cancellation_scopes)
        keywords = {"batch": batch}
        if on_reserved is not None:
            keywords["on_reserved"] = on_reserved
        reservation = original_prepare(self, request, **keywords)
        trace("reserved", task_id=request.request_id, name=Path(request.input_refs[0].path).name)
        if mode == "prepare_return":
            trace("ready")
            wait_release()
        return reservation

    def register(self, request, *, batch=False, on_reserved=None):
        keywords = {"batch": batch}
        if on_reserved is not None:
            keywords["on_reserved"] = on_reserved
        reservation = original_register(self, request, **keywords)
        if mode == "deadline_return":
            trace("deadline-return", task_id=request.request_id)
            trace("ready")
            wait_release()
        return reservation

    def cancel(self, task_id, *, reservation=None):
        if reservation is None:
            original_cancel(self, task_id)
        else:
            original_cancel(self, task_id, reservation=reservation)
        trace("cancelled-through-application", task_id=task_id)

    def stop(self):
        if mode in reservation_windows:
            before = list(self._cancellation_scopes)
            original = next(iter(requests_by_id.values()))
            reused_parent = root / "reused"
            reused_parent.mkdir()
            reused = replace(
                original,
                output_policy=replace(original.output_policy, output_dir=str(reused_parent), output_path=None),
            )
            try:
                result = original_execute(self, reused)
                trace(
                    "controller-reused",
                    before=before,
                    after=list(self._cancellation_scopes),
                    success=result.success,
                    outputs=[artifact.staging_path for artifact in result.artifacts if artifact.is_primary],
                    task_id=reused.request_id,
                )
            except Exception as exc:
                trace("controller-reuse-failed", before=before, error=repr(exc))
        original_stop(self)

    def execute(self, request):
        trace("worker", task_id=request.request_id, name=Path(request.input_refs[0].path).name)
        return original_execute(self, request)

    def pool_init(self, *args, **kwargs):
        caller = sys._getframe(1).f_globals.get("__name__", "")
        original_pool_init(self, *args, **kwargs)
        self._docwen_phase_pool = caller in {"docwen_cli.commands.convert", "docwen_cli.batch_execution"}
        if self._docwen_phase_pool and mode in {"before_worker", "submit_after_enqueue"}:
            # Real queued futures, not merely a wrapper already running on a
            # thread. Owned blockers are released by the parent in finally.
            for _ in range(self._max_workers):
                original_submit(self, wait_release)

    def submit(self, fn, /, *args, **kwargs):
        nonlocal submitted_count
        if not getattr(self, "_docwen_phase_pool", False):
            return original_submit(self, fn, *args, **kwargs)
        future = original_submit(self, fn, *args, **kwargs)
        submitted_count += 1
        if mode == "before_worker" and submitted_count == 2:
            trace("ready")
        elif mode in {"submit_after_enqueue", "submit_after_commit"} and submitted_count == 1:
            if mode == "submit_after_commit":
                future.result(timeout=6)
            trace("submit-window", done=future.done())
            trace("ready")
            wait_release()
        elif mode == "after_submit_return" and submitted_count == 1:
            # Stop at the caller's next line after submit returned. This is the
            # old future-assigned/before-map-insertion window, without relying
            # on a particular line number or source-line spelling.
            caller = sys._getframe(1)

            def after_return(frame, event, _arg):
                if frame is caller and event == "line":
                    sys.settrace(None)
                    caller.f_trace = None
                    future.result(timeout=6)
                    trace("submit-return-window", done=future.done())
                    trace("ready")
                    wait_release()
                return after_return

            caller.f_trace = after_return
            sys.settrace(after_return)
        return future

    def convert(self, context):
        nonlocal ready_count
        name = Path(context.request.input_refs[0].path).name
        result = original_convert(self, context)
        trace("staged", name=name, success=result.success)
        if mode == "during_conversion" or (mode in {"partial_interrupt", "partial_timeout"} and name == "second.csv"):
            with trace_lock:
                ready_count += 1
                ready = ready_count == (2 if mode == "during_conversion" else 1)
            if ready:
                trace("ready")
            if mode == "partial_timeout":
                wait_cancellation(context.cancellation)
            else:
                wait_release(context.cancellation)
        return result

    def finalize(self, *args, **kwargs):
        if mode == "before_publish":
            trace("ready")
            wait_cancellation(kwargs["cancellation"])
        result = original_finalize(self, *args, **kwargs)
        trace(
            "finalized",
            name=Path(kwargs.get("input_path", "")).name,
            success=result.success,
            outputs=[artifact.staging_path for artifact in result.artifacts],
        )
        if mode == "after_publish":
            trace("ready")
            wait_cancellation(kwargs["cancellation"])
        return result

    ApplicationController.prepare_execution_cancellation = prepare
    ApplicationController.cancel = cancel
    ApplicationController.execute_single = execute
    ApplicationController.stop = stop
    ExecutionDeadline.register = register
    CsvToXlsxConverter.convert = convert
    OutputFinalizer.finalize = finalize
    ThreadPoolExecutor.__init__ = pool_init
    ThreadPoolExecutor.submit = submit

    names = ["first.csv", "second.csv", "third.csv"]
    windows = {
        "before_worker",
        "during_conversion",
        "submit_after_enqueue",
        "submit_after_commit",
        "after_submit_return",
        *reservation_windows,
    }
    if mode in windows:
        names = names[:2]
    elif mode in {"before_publish", "after_publish"}:
        names = names[:1]
    jobs = "2" if mode in windows else "1"
    timeout = "2" if mode in {"partial_timeout", "before_publish", "after_publish"} else "10"
    return cli_main(
        [
            "batch",
            "convert",
            *(str(root / name) for name in names),
            "--to",
            "xlsx",
            "--output-dir",
            str(root / "published"),
            "--jobs",
            jobs,
            "--timeout",
            timeout,
            "--json",
            "--continue-on-error",
        ]
    )


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1], Path(sys.argv[2])))
