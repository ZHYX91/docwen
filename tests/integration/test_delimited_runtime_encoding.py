"""Real CSV/TSV encoding admission through SpreadsheetPlugin, Runtime, and final publication."""

import codecs
import csv
import io
from pathlib import Path

import openpyxl
import pytest

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def _encode_delimited_payload(text: str, encoding: str) -> bytes:
    prefix, codec = {
        "utf8-bom": (codecs.BOM_UTF8, "utf-8"),
        "utf16-le": (codecs.BOM_UTF16_LE, "utf-16-le"),
        "utf16-be": (codecs.BOM_UTF16_BE, "utf-16-be"),
        "utf32-le": (codecs.BOM_UTF32_LE, "utf-32-le"),
        "utf32-be": (codecs.BOM_UTF32_BE, "utf-32-be"),
        "gbk": (b"", "gbk"),
    }[encoding]
    return prefix + text.encode(codec)


@pytest.mark.integration
@pytest.mark.parametrize("source_format", ["csv", "tsv"])
@pytest.mark.parametrize(
    "encoding",
    ["utf8-bom", "utf16-le", "utf16-be", "utf32-le", "utf32-be", "gbk"],
)
@pytest.mark.parametrize(
    ("case_name", "expected_row"),
    [
        ("single-record", ["北京", "00123"]),
        ("quoted-multiline", ["上海", "第一行\n第二行"]),
    ],
)
def test_delimited_runtime_pipeline_honors_shared_encoding_contract(
    tmp_path: Path,
    source_format: str,
    encoding: str,
    case_name: str,
    expected_row: list[str],
) -> None:
    from docwen_core.detection import inspect_file
    from docwen_core.models import (
        FILE_INSPECTION_METADATA_KEY,
        ConversionRequest,
        FileRef,
        OutputPolicy,
    )
    from docwen_plugin_spreadsheet.plugin import SpreadsheetPlugin
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    separator = "," if source_format == "csv" else "\t"
    buffer = io.StringIO()
    csv.writer(buffer, delimiter=separator, lineterminator="\n").writerow(expected_row)

    source = tmp_path / f"{case_name}-{encoding}.{source_format}"
    source.write_bytes(_encode_delimited_payload(buffer.getvalue(), encoding))
    inspection = inspect_file(str(source))
    assert inspection.detected_format == source_format
    assert inspection.workflow_category == "spreadsheet"
    assert inspection.may_execute and not inspection.requires_explicit_acceptance

    output_dir = tmp_path / "published"
    workspace_root = tmp_path / "workspace"
    registry = PluginRegistry()
    registry.register(SpreadsheetPlugin())
    manager = TaskManager(
        registry,
        RouteResolver(registry),
        WorkspaceManager(root_dir=str(workspace_root)),
        OutputFinalizer(),
    )
    request = ConversionRequest(
        request_id=f"delimited-{source_format}-{encoding}-{case_name}",
        input_refs=[
            FileRef(
                path=str(source),
                format=inspection.detected_format,
                category=inspection.workflow_category,
                size_bytes=source.stat().st_size,
                metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
            )
        ],
        target_format="xlsx",
        output_policy=OutputPolicy(output_dir=str(output_dir)),
    )

    result = manager.execute_single(request)

    assert result.success is True, result.diagnostics
    assert result.error is None
    assert len(result.artifacts) == 1
    published = Path(result.artifacts[0].staging_path)
    assert published.parent == output_dir
    workbook = openpyxl.load_workbook(published)
    try:
        sheet = workbook.active
        assert sheet is not None
        assert [cell.value for cell in sheet[1]] == expected_row
        assert all(cell.data_type == "s" for cell in sheet[1])
    finally:
        workbook.close()
    assert not list(workspace_root.rglob("*.xlsx"))
