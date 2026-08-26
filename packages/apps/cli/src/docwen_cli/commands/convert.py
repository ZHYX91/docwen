"""Application-facing execution service for protocol 3 domain commands.

Public command modules normalize their explicit command trees into one
application request shape here. Runtime action identifiers remain internal and
are never accepted as compatibility aliases at the parser boundary.
"""

from __future__ import annotations

import argparse
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.runtime_capability_catalog import RuntimeCapabilityCatalog, RuntimeRoute
from docwen_cli.capabilities import normalize_target_format, runtime_route_catalog
from docwen_cli.commands import execution_request
from docwen_cli.commands.execution_errors import (
    print_invalid_input as _print_invalid_input,
)
from docwen_cli.commands.execution_errors import (
    print_unavailable as _print_unavailable,
)
from docwen_cli.commands.execution_options import build_execution_options, validate_execution_options
from docwen_cli.execution_deadline import Deadline, ExecutionDeadline, NoopDeadline, timeout_result
from docwen_cli.exit_codes import ExitCode, exit_code_from_error_code
from docwen_cli.file_admission_i18n import (
    detected_format_acceptance_hint,
    render_file_inspection_message,
    render_file_inspection_warning,
)
from docwen_cli.i18n import cli_t
from docwen_cli.utils import (
    create_progress_callback,
    expand_paths,
    validate_files,
)
from docwen_core.detection import admission_error_type, inspect_file
from docwen_core.errors import DocWenError
from docwen_core.models.file_inspection import FileInspection
from docwen_core.models.request import ConversionRequest, FileRef

# ── Execution through application layer ────────────────────────────


def _resolve_runtime_routes(
    catalog: RuntimeCapabilityCatalog,
    *,
    files: list[str],
    inspections: dict[str, FileInspection],
    action: str,
    target_format: str,
) -> dict[str, RuntimeRoute]:
    """Resolve every admitted file through the canonical route projection."""

    resolved: dict[str, RuntimeRoute] = {}
    for file_path in files:
        inspection = _inspection_for(file_path, inspections)
        if action and not target_format:
            route = catalog.resolve_action_route(
                detected_format=inspection.detected_format,
                workflow_category=inspection.workflow_category,
                action_name=action,
            )
        else:
            if not target_format:
                raise ValueError("Conversion target must be explicit.")
            route = catalog.resolve_route(
                detected_format=inspection.detected_format,
                workflow_category=inspection.workflow_category,
                action_name=action,
                target=target_format,
            )
        if route is None:
            requested_operation = "action" if action else "conversion"
            valid_targets = sorted(
                {
                    candidate.target
                    for candidate in catalog.effective_routes(
                        inspection.detected_format,
                        inspection.workflow_category,
                    )
                    if candidate.available
                    and candidate.operation == requested_operation
                    and candidate.action_name == action
                }
            )
            target_hint = f" Available targets: {', '.join(valid_targets)}." if valid_targets else ""
            requested_target = target_format or "<runtime-owned>"
            raise ValueError(
                "No canonical runtime route for "
                f"{inspection.detected_format} -> {requested_target} (action={action!r})."
                f"{target_hint}"
            )
        if not route.available:
            raise CapabilityUnavailableError(f"Runtime route is unavailable: {route.id} ({route.state}).")
        resolved[file_path] = route
    return resolved


def execute_convert(
    args: argparse.Namespace,
    controller: Any | None = None,
) -> int:
    """Execute a normalized protocol 3 domain command.

    Args:
        args: Parsed argparse namespace.
        controller: ``ApplicationController`` instance (conversion always
            needs a controller).

    Returns:
        Exit code.
    """
    # Resolve action from --action flag; default to "" (standard conversion).
    action = execution_request.resolve_cli_action(args)

    # Conversion targets are explicit. Named-action targets are resolved from
    # the canonical Runtime route after content admission.
    target_format = normalize_target_format(getattr(args, "to", ""))
    if target_format:
        args.to = target_format

    # Resolve files — handle both "files" (list) and "file" (single) args
    raw_files: list[str] = []
    if hasattr(args, "files") and args.files:
        raw_files = list(args.files)
    elif hasattr(args, "file") and args.file:
        raw_files = [str(args.file)]
    files = expand_paths(raw_files)

    if not files:
        return _print_invalid_input(
            action,
            args,
            cli_t("cli.messages.error_file_not_found"),
        )

    inspections: dict[str, FileInspection] = {}
    valid_files, invalid_files, admission_messages = validate_files(
        files,
        use_detected_format=bool(getattr(args, "use_detected_format", False)),
        inspection_cache=inspections,
    )

    # Human mode may summarize invalid inputs. Machine mode carries the same
    # information in its typed result and must not emit a second stream.
    if invalid_files and not getattr(args, "quiet", False) and not getattr(args, "json", False):
        out = sys.stdout
        print(
            cli_t("cli.messages.warning_invalid_files", count=len(invalid_files)),
            file=out,
        )
        for file, reason in invalid_files:
            print(f"  - {file}: {reason}", file=out)

    if admission_messages and not getattr(args, "quiet", False) and not getattr(args, "json", False):
        for file, message in admission_messages:
            print(f"WARNING: {file}: {message}", file=sys.stderr)

    if not valid_files:
        inspection = inspections.get(os.path.abspath(files[0])) if len(files) == 1 else None
        if inspection is not None:
            return _print_invalid_input(
                action,
                args,
                render_file_inspection_message(inspection),
                input_file=inspection.file_path,
                error_code=_admission_error_code(inspection),
                details={"file": inspection.file_path, "admission": inspection.to_dict()},
                hint=(detected_format_acceptance_hint() if inspection.requires_explicit_acceptance else None),
            )
        return _print_invalid_input(
            action,
            args,
            cli_t("cli.messages.error_no_valid_files"),
        )

    try:
        validate_execution_options(args)
    except ValueError as e:
        return _print_invalid_input(action, args, str(e))

    if controller is None or not getattr(controller, "has_runtime", False):
        return _print_unavailable(action, args, "Runtime 未初始化，无法解析或执行转换")
    try:
        route_catalog = runtime_route_catalog(controller)
        routes_by_file = _resolve_runtime_routes(
            route_catalog,
            files=valid_files,
            inspections=inspections,
            action=action,
            target_format=target_format,
        )
    except ValueError as exc:
        return _print_invalid_input(action, args, str(exc))
    except CapabilityUnavailableError as exc:
        return _print_invalid_input(action, args, str(exc), error_code="capability_unavailable")

    common_route_options = set.intersection(*(set(route.options) for route in routes_by_file.values()))
    try:
        options = build_execution_options(
            args,
            route_options=common_route_options,
        )
        needs_ocr_language = bool(getattr(args, "ocr", False)) and any(
            "ocr_language" in route.options for route in routes_by_file.values()
        )
        configured_ocr_language = execution_request.configured_ocr_language(controller) if needs_ocr_language else None
    except ValueError as exc:
        return _print_invalid_input(action, args, str(exc))
    except CapabilityUnavailableError as exc:
        return _print_invalid_input(action, args, str(exc), error_code="capability_unavailable")

    # ── dry-run path ─────────────────────────────────────────────
    if getattr(args, "dry_run", False):
        if len(valid_files) != 1:
            return _print_invalid_input(
                action,
                args,
                "--dry-run 当前仅支持单文件请求",
                input_file=valid_files[0] if valid_files else "",
            )
        try:
            return _execute_dry_run(
                action,
                valid_files[0],
                target_format,
                options,
                args,
                inspection=_inspection_for(valid_files[0], inspections),
                route=routes_by_file[valid_files[0]],
                configured_ocr_language=configured_ocr_language,
            )
        except ValueError as exc:
            return _print_invalid_input(action, args, str(exc))

    # ── Convert path ─────────────────────────────────────────────
    json_mode = bool(getattr(args, "json", False))
    include_timing = bool(getattr(args, "timing", False))
    continue_on_error = bool(getattr(args, "continue_on_error", False))
    is_batch = len(valid_files) > 1 or bool(getattr(args, "batch", False))
    jobs = max(1, int(getattr(args, "jobs", 1) or 1))
    json_invalid_files = invalid_files if json_mode and is_batch else []
    json_input_files = files if json_invalid_files else None

    progress_cb = create_progress_callback(
        quiet=getattr(args, "quiet", False),
        verbose=getattr(args, "verbose", False),
        json_mode=json_mode,
    )
    deadline = ExecutionDeadline(controller, float(getattr(args, "timeout", 600))).start()

    try:
        # ── Aggregate path (merge-* actions: many-to-one) ──────────
        if execution_request.public_command(args).startswith("merge "):
            if len(valid_files) < 2:
                return _print_invalid_input(
                    action,
                    args,
                    cli_t("cli.messages.error_aggregate_need_two", count=len(valid_files)),
                )
            return _execute_aggregate(
                controller,
                action,
                valid_files,
                target_format,
                options,
                args,
                json_mode=json_mode,
                include_timing=include_timing,
                progress_cb=progress_cb,
                deadline=deadline,
                inspections=inspections,
                routes_by_file=routes_by_file,
                configured_ocr_language=configured_ocr_language,
            )

        if not is_batch:
            return _execute_single(
                controller,
                action,
                valid_files[0],
                target_format,
                options,
                args,
                json_mode=json_mode,
                include_timing=include_timing,
                progress_cb=progress_cb,
                deadline=deadline,
                inspection=_inspection_for(valid_files[0], inspections),
                route=routes_by_file[valid_files[0]],
                configured_ocr_language=configured_ocr_language,
            )
        else:
            return _execute_batch(
                controller,
                action,
                valid_files,
                target_format,
                options,
                args,
                json_mode=json_mode,
                include_timing=include_timing,
                continue_on_error=continue_on_error,
                max_workers=jobs,
                progress_cb=progress_cb,
                invalid_files=json_invalid_files,
                original_input_files=json_input_files,
                deadline=deadline,
                inspections=inspections,
                routes_by_file=routes_by_file,
                configured_ocr_language=configured_ocr_language,
            )
    except ValueError as exc:
        return _print_invalid_input(action, args, str(exc))
    except KeyboardInterrupt:
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            presenter = JsonPresenter(include_timing=include_timing)
            presenter.present_error(
                execution_request.public_command(args), "用户中断", error_code="operation_cancelled"
            )
        else:
            print("错误: 用户中断", file=sys.stderr)
        return int(ExitCode.CANCELLED)
    finally:
        deadline.close()


# ── Internal helpers ───────────────────────────────────────────────


def _inspection_for(file_path: str, inspections: dict[str, FileInspection]) -> FileInspection:
    key = os.path.abspath(file_path)
    inspection = inspections.get(key)
    if inspection is None:
        inspection = inspect_file(key)
        inspections[key] = inspection
    return inspection


def _admission_error_code(inspection: FileInspection) -> str:
    return admission_error_type(inspection)


def _admission_warning_payloads(*inspections: FileInspection) -> list[dict[str, Any]]:
    payloads: list[dict[str, Any]] = []
    for inspection in inspections:
        if not inspection.warning_code:
            continue
        payloads.append(
            {
                "level": "warning",
                "code": inspection.warning_code,
                "message": render_file_inspection_warning(inspection),
                "file": inspection.file_path,
                "details": {
                    "relation": inspection.relation.value,
                    "decision": inspection.decision.value,
                    "declared_format": inspection.declared_format,
                    "detected_format": inspection.detected_format,
                },
            }
        )
    return payloads


def _add_json_warnings(presenter: Any, inspections: list[FileInspection]) -> None:
    for warning in _admission_warning_payloads(*inspections):
        presenter.add_warning(warning)


def _execute_single(
    controller: Any,
    action: str,
    file_path: str,
    target_format: str,
    options: dict[str, Any],
    args: argparse.Namespace,
    *,
    json_mode: bool = False,
    include_timing: bool = False,
    progress_cb: Any = None,
    deadline: Deadline | None = None,
    inspection: FileInspection | None = None,
    route: RuntimeRoute,
    configured_ocr_language: str | None = None,
) -> int:
    """Execute a single-file conversion through ApplicationController."""
    import uuid

    inspection = inspection or inspect_file(file_path)
    request_action, request_target = route.action_name, route.target
    input_ref = execution_request.file_ref_for_runtime(
        file_path,
        inspection,
        explicit_acceptance=bool(getattr(args, "use_detected_format", False)),
    )
    request_options = execution_request.project_route_options(
        options,
        route_id=route.id,
        route_options=route.options,
        configured_ocr_language=configured_ocr_language,
        ocr_requested=bool(getattr(args, "ocr", False)),
    )

    request = ConversionRequest(
        request_id=str(uuid.uuid4()),
        input_refs=[input_ref],
        target_format=request_target,
        action_name=request_action,
        options=request_options,
        output_policy=execution_request.output_policy(args),
    )

    deadline = deadline or NoopDeadline()
    reservation = deadline.register(request)
    try:
        result = controller.execute_single(request)
    except FileNotFoundError as exc:
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            presenter = JsonPresenter(include_timing=include_timing)
            _add_json_warnings(presenter, [inspection])
            presenter.present_error(
                execution_request.public_command(args),
                str(exc),
                error_code="invalid_input",
            )
        else:
            print(f"错误: {exc}", file=sys.stderr)
        return int(ExitCode.INVALID_INPUT)
    except DocWenError as exc:
        error_code = getattr(exc, "error_type", "conversion_failed")
        details = getattr(exc, "details", None)
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            presenter = JsonPresenter(include_timing=include_timing)
            _add_json_warnings(presenter, [inspection])
            presenter.present_error(
                execution_request.public_command(args),
                str(exc),
                error_code=str(error_code),
                details=details,
            )
        else:
            msg = str(exc)
            if request_action:
                msg += (
                    f"\n提示: 当前 source/target/action 组合 '{request_action}' 无匹配插件支持。"
                    " 请运行 `docwen resources list optimizations` 查看可用动作，"
                    "或运行 `docwen inspect <file>` 查看文件支持的操作。"
                )
            print(f"错误: {msg}", file=sys.stderr)
        return int(exit_code_from_error_code(str(error_code)))
    except Exception as exc:
        error_code = str(getattr(exc, "code", "conversion_failed"))
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            presenter = JsonPresenter(include_timing=include_timing)
            _add_json_warnings(presenter, [inspection])
            presenter.present_error(execution_request.public_command(args), str(exc), error_code=error_code)
        else:
            print(f"错误: {exc}", file=sys.stderr)
        return int(exit_code_from_error_code(error_code))
    finally:
        deadline.release(request, reservation)

    deadline.finish()
    result = timeout_result(result, timed_out=deadline.timed_out)

    if json_mode:
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter(include_timing=include_timing)
        _add_json_warnings(presenter, [inspection])
        presenter.present_single(
            result,
            command=execution_request.public_command(args),
            action_name=request_action,
            input_files=[file_path],
        )
    else:
        from docwen_cli.presenters.text_presenter import TextPresenter

        presenter = TextPresenter(
            quiet=getattr(args, "quiet", False),
            verbose=getattr(args, "verbose", False),
        )
        presenter.present_single(result)

    if progress_cb:
        progress_cb("完成")

    return (
        int(ExitCode.OK)
        if getattr(result, "success", False)
        else int(exit_code_from_error_code(getattr(getattr(result, "error", None), "error_type", "unknown_error")))
    )


def _execute_batch(
    controller: Any,
    action: str,
    files: list[str],
    target_format: str,
    options: dict[str, Any],
    args: argparse.Namespace,
    *,
    json_mode: bool = False,
    include_timing: bool = False,
    continue_on_error: bool = False,
    max_workers: int = 1,
    progress_cb: Any = None,
    invalid_files: list[tuple[str, str]] | None = None,
    original_input_files: list[str] | None = None,
    deadline: Deadline | None = None,
    inspections: dict[str, FileInspection] | None = None,
    routes_by_file: dict[str, RuntimeRoute],
    configured_ocr_language: str | None = None,
) -> int:
    """Execute a batch conversion using ``ThreadPoolExecutor`` with *max_workers*.

    Each file is converted in its own single-file request via
    ``controller.execute_single()``, respecting the ``--jobs`` parameter.
    """
    import uuid

    deadline = deadline or NoopDeadline()
    inspections = inspections if inspections is not None else {}

    # Bound concurrency to protect Office providers and memory-heavy converters.
    max_workers = min(max_workers, min(4, (os.cpu_count() or 2)))
    if max_workers < 1:
        max_workers = 1

    invalid_by_file = dict(invalid_files or [])
    presentation_files = list(original_input_files or files)
    results: list[Any] = [None] * len(presentation_files)
    files_to_convert: list[tuple[int, str]] = []
    for index, file_path in enumerate(presentation_files):
        invalid_reason = invalid_by_file.get(file_path)
        if invalid_reason is not None:
            results[index] = _invalid_file_result(index, invalid_reason)
        else:
            files_to_convert.append((index, file_path))
    seen_result_count = 0
    stop_on_error = not continue_on_error

    def _build_request(file_path: str) -> ConversionRequest:
        inspection = _inspection_for(file_path, inspections)
        route = routes_by_file[file_path]
        request_action, request_target = route.action_name, route.target
        input_ref = execution_request.file_ref_for_runtime(
            file_path,
            inspection,
            explicit_acceptance=bool(getattr(args, "use_detected_format", False)),
        )
        request_options = execution_request.project_route_options(
            options,
            route_id=route.id,
            route_options=route.options,
            configured_ocr_language=configured_ocr_language,
            ocr_requested=bool(getattr(args, "ocr", False)),
        )
        return ConversionRequest(
            request_id=str(uuid.uuid4()),
            input_refs=[input_ref],
            target_format=request_target,
            action_name=request_action,
            options=request_options,
            output_policy=execution_request.output_policy(args),
        )

    requests_to_convert = [(result_index, _build_request(file_path)) for result_index, file_path in files_to_convert]
    reservations = {
        str(request.request_id): deadline.register(request) for _result_index, request in requests_to_convert
    }

    def _convert_one(request: ConversionRequest) -> Any:
        try:
            return controller.execute_single(request)
        finally:
            deadline.release(request, reservations[str(request.request_id)])

    try:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all files
            future_to_index: dict[Any, int] = {}
            for submitted_index, (result_index, request) in enumerate(requests_to_convert):
                if progress_cb:
                    progress_cb(f"... {submitted_index + 1}/{len(files_to_convert)}")
                future = executor.submit(_convert_one, request)
                future_to_index[future] = result_index

            # Collect results in original order
            for future in as_completed(future_to_index):
                idx = future_to_index[future]
                try:
                    result = future.result()
                except Exception as exc:
                    from docwen_core.models.result import (
                        ConversionErrorInfo,
                        ConversionMetrics,
                        ConversionResult,
                    )

                    result = ConversionResult(
                        task_id=f"batch-{idx}",
                        success=False,
                        error=ConversionErrorInfo(
                            error_type="conversion_failed",
                            message=str(exc),
                        ),
                        metrics=ConversionMetrics(),
                    )

                results[idx] = result
                seen_result_count += 1

                if progress_cb:
                    progress_cb(f"... {seen_result_count}/{len(files)}")

                # Stop on error: cancel remaining futures
                if stop_on_error and not getattr(result, "success", False):
                    for f in future_to_index:
                        if not f.done():
                            f.cancel()
                    # Mark remaining as skipped
                    for j, r in enumerate(results):
                        if r is None and j > idx:
                            from docwen_core.models.result import (
                                ConversionErrorInfo,
                                ConversionMetrics,
                                ConversionResult,
                            )

                            results[j] = ConversionResult(
                                task_id=f"batch-{j}",
                                success=False,
                                error=ConversionErrorInfo(
                                    error_type="skipped",
                                    message="Skipped due to previous error",
                                ),
                                metrics=ConversionMetrics(),
                            )
                    break
    except KeyboardInterrupt:
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            p = JsonPresenter(include_timing=include_timing)
            _add_json_warnings(p, [_inspection_for(path, inspections) for path in files])
            p.present_error(execution_request.public_command(args), "用户中断", error_code="operation_cancelled")
        else:
            print("错误: 用户中断", file=sys.stderr)
        return int(ExitCode.CANCELLED)

    deadline.finish()

    # Fill in any None slots (cancelled/skipped)
    for i in range(len(results)):
        if results[i] is None:
            from docwen_core.models.result import (
                ConversionErrorInfo,
                ConversionMetrics,
                ConversionResult,
            )

            results[i] = ConversionResult(
                task_id=f"batch-{i}",
                success=False,
                error=ConversionErrorInfo(
                    error_type="skipped",
                    message="Skipped",
                ),
                metrics=ConversionMetrics(),
            )

    results = [timeout_result(result, timed_out=deadline.timed_out) for result in results]

    if json_mode:
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter(include_timing=include_timing)
        _add_json_warnings(presenter, [_inspection_for(path, inspections) for path in files])
        presenter.present_batch(
            results,
            command=execution_request.public_command(args),
            action_name=action,
            input_files=presentation_files,
        )
    else:
        from docwen_cli.presenters.text_presenter import TextPresenter

        presenter = TextPresenter(
            quiet=getattr(args, "quiet", False),
            verbose=getattr(args, "verbose", False),
        )
        presenter.present_batch(results, input_files=presentation_files)

    if progress_cb:
        progress_cb("完成")

    return _batch_exit_code(results)


def _batch_exit_code(results: list[Any]) -> int:
    """Return the stable aggregate exit code for a completed batch."""
    success_count = sum(1 for result in results if getattr(result, "success", False))
    failed_count = len(results) - success_count
    if failed_count == 0:
        return int(ExitCode.OK)
    if success_count:
        return int(exit_code_from_error_code("batch_partial_failure"))

    first_error_type = "unknown_error"
    for result in results:
        if getattr(result, "success", False):
            continue
        first_error_type = str(getattr(getattr(result, "error", None), "error_type", "unknown_error"))
        break
    return int(exit_code_from_error_code(first_error_type))


def _invalid_file_result(index: int, reason: str) -> Any:
    from docwen_core.models.result import ConversionErrorInfo, ConversionMetrics, ConversionResult

    return ConversionResult(
        task_id=f"batch-invalid-{index}",
        success=False,
        error=ConversionErrorInfo(
            error_type="invalid_input",
            message=reason,
            diagnostic_code="unsupported_extension" if "扩展名" in reason or "extension" in reason else "invalid_input",
        ),
        metrics=ConversionMetrics(),
    )


def _execute_aggregate(
    controller: Any,
    action: str,
    files: list[str],
    target_format: str,
    options: dict[str, Any],
    args: argparse.Namespace,
    *,
    json_mode: bool = False,
    include_timing: bool = False,
    progress_cb: Any = None,
    deadline: Deadline | None = None,
    inspections: dict[str, FileInspection] | None = None,
    routes_by_file: dict[str, RuntimeRoute],
    configured_ocr_language: str | None = None,
) -> int:
    """Execute an aggregate (merge) operation — many files → one output.

    All input files are passed together in a single ``ConversionRequest``
    with multiple ``input_refs``.  The application layer's
    ``AggregateWorkflow`` passes them to the runtime as a group so the
    merge-capable converter produces a single merged output.
    """
    import uuid

    inspections = inspections if inspections is not None else {}

    if len(files) < 2:
        return _print_invalid_input(
            action,
            args,
            cli_t("cli.messages.error_aggregate_need_two", count=len(files)),
        )

    route = routes_by_file[files[0]]
    if any(routes_by_file[file_path].id != route.id for file_path in files[1:]):
        return _print_invalid_input(
            action,
            args,
            "Aggregate inputs do not resolve to one canonical runtime route.",
        )
    request_action, request_target = route.action_name, route.target

    input_refs: list[FileRef] = []
    for fp in files:
        inspection = _inspection_for(fp, inspections)
        input_refs.append(
            execution_request.file_ref_for_runtime(
                fp,
                inspection,
                explicit_acceptance=bool(getattr(args, "use_detected_format", False)),
            )
        )
    request_options = execution_request.project_route_options(
        options,
        route_id=route.id,
        route_options=route.options,
        configured_ocr_language=configured_ocr_language,
        ocr_requested=bool(getattr(args, "ocr", False)),
    )

    request = ConversionRequest(
        request_id=str(uuid.uuid4()),
        input_refs=input_refs,
        target_format=request_target,
        action_name=request_action,
        options=request_options,
        output_policy=execution_request.output_policy(args),
    )

    deadline = deadline or NoopDeadline()
    reservation = deadline.register(request)
    try:
        result = controller.execute_aggregate(request, action)
    except DocWenError as exc:
        error_code = getattr(exc, "error_type", "conversion_failed")
        details = getattr(exc, "details", None)
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            presenter = JsonPresenter(include_timing=include_timing)
            _add_json_warnings(presenter, [_inspection_for(path, inspections) for path in files])
            presenter.present_error(
                execution_request.public_command(args),
                str(exc),
                error_code=str(error_code),
                details=details,
            )
        else:
            msg = str(exc)
            if action:
                msg += (
                    f"\n提示: 当前 source/target/action 组合 '{action}' 无匹配插件支持。"
                    " 请运行 `docwen resources list optimizations` 查看可用动作，"
                    "或运行 `docwen inspect <file>` 查看文件支持的操作。"
                )
            print(f"错误: {msg}", file=sys.stderr)
        return int(exit_code_from_error_code(str(error_code)))
    except Exception as exc:
        error_code = str(getattr(exc, "code", "conversion_failed"))
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            presenter = JsonPresenter(include_timing=include_timing)
            _add_json_warnings(presenter, [_inspection_for(path, inspections) for path in files])
            presenter.present_error(execution_request.public_command(args), str(exc), error_code=error_code)
        else:
            print(f"错误: {exc}", file=sys.stderr)
        return int(exit_code_from_error_code(error_code))
    finally:
        deadline.release(request, reservation)

    deadline.finish()
    result = timeout_result(result, timed_out=deadline.timed_out)

    if json_mode:
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter(include_timing=include_timing)
        _add_json_warnings(presenter, [_inspection_for(path, inspections) for path in files])
        presenter.present_single(
            result,
            command=execution_request.public_command(args),
            action_name=action,
            input_files=files,
        )
    else:
        from docwen_cli.presenters.text_presenter import TextPresenter

        presenter = TextPresenter(
            quiet=getattr(args, "quiet", False),
            verbose=getattr(args, "verbose", False),
        )
        presenter.present_single(result)

    if progress_cb:
        progress_cb("完成")

    return (
        int(ExitCode.OK)
        if getattr(result, "success", False)
        else int(exit_code_from_error_code(getattr(getattr(result, "error", None), "error_type", "unknown_error")))
    )


def _execute_dry_run(
    action: str,
    file_path: str,
    target_format: str,
    options: dict[str, Any],
    args: argparse.Namespace,
    *,
    inspection: FileInspection | None = None,
    route: RuntimeRoute,
    configured_ocr_language: str | None = None,
) -> int:
    """Execute a dry-run: detect format, resolve route, print info."""
    json_mode = bool(getattr(args, "json", False))
    inspection = inspection or inspect_file(file_path)
    request_action, request_target = route.action_name, route.target
    route_ref = execution_request.file_ref_for_runtime(
        file_path,
        inspection,
        explicit_acceptance=bool(getattr(args, "use_detected_format", False)),
    )
    source_format = str(route_ref.format or "unknown").strip().lower()
    source_category = str(route_ref.category or "other").strip().lower()
    source_candidates = list(
        dict.fromkeys(
            candidate
            for candidate in (source_format, source_category)
            if candidate and candidate not in {"unknown", "other"}
        )
    )
    category_fallback = source_category if source_category not in {"", "other", source_format} else None
    effective_options = execution_request.project_route_options(
        options,
        route_id=route.id,
        route_options=route.options,
        configured_ocr_language=configured_ocr_language,
        ocr_requested=bool(getattr(args, "ocr", False)),
    )

    if json_mode:
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter()
        _add_json_warnings(presenter, [inspection])
        presenter.present_data(
            command=execution_request.public_command(args),
            data={
                "detected_format": inspection.detected_format,
                "detected_category": inspection.detected_category,
                "workflow_category": inspection.workflow_category,
                "routing": {
                    "status": "deferred_to_runtime",
                    "source_format": source_format,
                    "source_category": source_category,
                    "source_candidates": source_candidates,
                    "category_fallback": category_fallback,
                    "target_format": request_target,
                    "action_name": request_action,
                },
                "expected_output": f"{request_target}",
                "effective_options": execution_request.redacted_options(effective_options),
                "admission": inspection.to_dict(),
            },
            success=True,
        )
    else:
        print(f"格式: {inspection.detected_format}")
        print(f"检测类别: {inspection.detected_category}")
        print(f"工作流类别: {inspection.workflow_category}")
        print(f"准入决定: {inspection.decision.value}")
        if inspection.warning_message:
            print(f"WARNING: {render_file_inspection_warning(inspection)}", file=sys.stderr)
        print(f"路由候选（按顺序）: {' -> '.join(source_candidates) or '无'}")
        if category_fallback:
            print(f"类别回退: {category_fallback}")
        print("路由解析: 执行时由 runtime 确定")
        print(f"目标输出: {request_target}")
        effective = execution_request.redacted_options(effective_options)
        if effective:
            print("有效选项:", effective)

    return int(ExitCode.OK)


# ── Registrar for argparser ────────────────────────────────────────
