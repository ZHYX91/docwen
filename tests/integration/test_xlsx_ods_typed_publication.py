"""Typed XLSX→ODS policy through the real Runtime, finalizer and Artifact Bundle.

Fixtures are newly generated, saved and reopened. Only the external Office
backend is substituted; the package policy, route, workspace and publication
are real. No Excel/LibreOffice execution or native-host acceptance is claimed.
"""

from __future__ import annotations

import hashlib
import shutil
from decimal import Decimal
from pathlib import Path
from typing import Any

import pytest

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.office_bridge import BridgeResult
from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
from tests.integration._xlsx_ods_typed_support import (
    FORMULA,
    ODS_MEDIA_TYPE,
    SourceCells,
    TypedSheets,
    assert_prepared_xlsx,
    ods_cell,
    ods_row,
    ods_typed_matrix,
    write_external_xlsx,
    write_ods,
    xlsx_cache_facts,
)

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def _source(path: Path, cells: SourceCells) -> bytes:
    original = write_external_xlsx(path, cells)
    assert xlsx_cache_facts(path) == {
        (sheet, coordinate): (kind, value is not None, value, FORMULA)
        for sheet, entries in cells.items()
        for coordinate, (kind, value) in entries.items()
    }
    return original


def _backend(monkeypatch: pytest.MonkeyPatch, delivery: Path, cells: SourceCells, prepared_copy: Path) -> None:
    def convert(input_path: str, output_path: str, **_kwargs: Any) -> BridgeResult:
        # Observe the independently decoded private backend input, not a policy helper.
        assert Path(input_path).suffix == ".xlsx"
        assert Path(output_path).suffix == ".ods"
        assert_prepared_xlsx(Path(input_path), cells)
        shutil.copyfile(input_path, prepared_copy)
        shutil.copyfile(delivery, output_path)
        return BridgeResult(True, output_path=output_path, backend="typed-ods-fixture")

    monkeypatch.setattr(sheet_module, "convert_with_backend_priority", convert)


def _request(source: Path, output: Path, request_id: str) -> ConversionRequest:
    return ConversionRequest(
        request_id=request_id,
        input_refs=[FileRef(path=str(source), format="xlsx", category="spreadsheet", size_bytes=source.stat().st_size)],
        target_format="ods",
        output_policy=OutputPolicy(output_dir=str(output)),
    )


@pytest.mark.parametrize(
    ("cells", "sheets", "expected"),
    [
        pytest.param(
            {"Data": {"A1": ("str", "hello")}},
            [("Data", ods_row(ods_cell("string", "hello")))],
            {"Data": [[("string", "hello")]]},
            id="F4-string-text-p",
        ),
        pytest.param(
            {"Data": {"A1": ("str", "")}},
            [("Data", ods_row(ods_cell("string", "")))],
            {"Data": [[("string", "")]]},
            id="F4-known-empty-string",
        ),
        pytest.param(
            {"Data": {"A1": ("str", "")}},
            [("Data", ods_row(ods_cell("blank")))],
            {"Data": [[("blank", None)]]},
            id="F4-known-empty-physical-blank",
        ),
        pytest.param(
            {"Data": {"A1": ("str", "hello")}},
            [("Data", ods_row('<table:table-cell office:value-type="string" office:string-value="hello"/>'))],
            {"Data": [[("string", "hello")]]},
            id="F4-string-value-attribute",
        ),
        pytest.param(
            {"Data": {"A1": ("n", "0"), "B1": ("n", "30")}},
            [("Data", ods_row(ods_cell("number", "0.0") + ods_cell("number", "3E1")))],
            {"Data": [[("number", Decimal(0)), ("number", Decimal(30))]]},
            id="F6-number-zero-and-equivalent-decimal",
        ),
        pytest.param(
            {"Data": {"A1": ("b", "0"), "B1": ("b", "1")}},
            [("Data", ods_row(ods_cell("boolean", "false") + ods_cell("boolean", "true")))],
            {"Data": [[("boolean", False), ("boolean", True)]]},
            id="F4-booleans-false-true",
        ),
        pytest.param(
            {"Data": {"A1": ("n", "30"), "B1": ("n", "30"), "A2": ("n", "30"), "B2": ("n", "30")}},
            [("Data", ods_row(ods_cell("number", "30", repeat=2), repeat=2))],
            {
                "Data": [
                    [("number", Decimal(30)), ("number", Decimal(30))],
                    [("number", Decimal(30)), ("number", Decimal(30))],
                ]
            },
            id="F6-row-column-repeat-positive",
        ),
        pytest.param(
            {"Data": {"B2": ("n", "30")}, "Second": {"A1": ("str", "hello")}},
            [
                ("Data", ods_row(ods_cell("blank", repeat=2)) + ods_row(ods_cell("blank") + ods_cell("number", "30"))),
                ("Second", ods_row(ods_cell("string", "hello"))),
            ],
            {
                "Data": [[("blank", None), ("blank", None)], [("blank", None), ("number", Decimal(30))]],
                "Second": [[("string", "hello")]],
            },
            id="F6-sheet-coordinate-identity-positive",
        ),
    ],
)
def test_typed_external_cache_delivery_matches_final_independent_matrix(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    round_trip_runtime: Any,
    cells: SourceCells,
    sheets: list[tuple[str, str]],
    expected: TypedSheets,
) -> None:
    source = tmp_path / "linked.xlsx"
    original = _source(source, cells)
    delivery = tmp_path / "backend-delivery.ods"
    delivered = write_ods(delivery, sheets)
    assert ods_typed_matrix(delivery) == expected
    prepared = tmp_path / "observed-private.xlsx"
    _backend(monkeypatch, delivery, cells, prepared)
    output = tmp_path / "published"
    request = _request(source, output, "typed-ods-positive")
    result = round_trip_runtime.execute(request)
    assert result.success, result.error
    assert_prepared_xlsx(prepared, cells)
    assert "EXTERNAL_LINK_FLATTENED" in {item.code for item in result.diagnostics}
    bundle = ArtifactBundleCommitter().commit(
        task_id=request.request_id,
        staging_root=str(output),
        draft=build_bundle_draft(
            profile="single_document", output_media_type=ODS_MEDIA_TYPE, artifacts=result.artifacts
        ),
    )
    assert len(bundle.entries) == 1
    main = next(item for item in bundle.artifacts if item.artifact_id == bundle.entries[0].artifact_id)
    published = output / main.locator
    assert published.read_bytes() == delivered
    assert main.sha256 == hashlib.sha256(delivered).hexdigest()
    assert ods_typed_matrix(published) == expected
    assert source.read_bytes() == original


@pytest.mark.parametrize(
    ("cells", "sheets", "expected"),
    [
        pytest.param(
            {"Data": {"A1": ("n", "30")}},
            [("Data", ods_row(ods_cell("number", "29") + ods_cell("number", "30")))],
            {"Data": [[("number", Decimal(29)), ("number", Decimal(30))]]},
            id="F6-A1-changed-value-bag-still-contains-30",
        ),
        pytest.param(
            {"Data": {"A1": ("n", "30")}},
            [("Data", ods_row(ods_cell("number", "29"))), ("WrongSheet", ods_row(ods_cell("number", "30")))],
            {"Data": [[("number", Decimal(29))]], "WrongSheet": [[("number", Decimal(30))]]},
            id="F6-value-moved-to-other-sheet",
        ),
        pytest.param(
            {"Data": {"A1": ("n", "30"), "B1": ("n", "30")}},
            [("Data", ods_row(ods_cell("number", "30")))],
            {"Data": [[("number", Decimal(30))]]},
            id="F6-one-cell-cannot-satisfy-two-identities",
        ),
        pytest.param(
            {"Data": {"A1": ("n", "30")}},
            [("Data", ods_row(ods_cell("string", "30")))],
            {"Data": [[("string", "30")]]},
            id="F6-number-changed-to-string",
        ),
        pytest.param(
            {"Data": {"A1": ("b", "0")}},
            [("Data", ods_row(ods_cell("number", "0")))],
            {"Data": [[("number", Decimal(0))]]},
            id="F6-false-changed-to-zero",
        ),
        pytest.param(
            {"Data": {"A1": ("str", "30")}},
            [("Data", ods_row(ods_cell("number", "30")))],
            {"Data": [[("number", Decimal(30))]]},
            id="F6-string-changed-to-number",
        ),
        pytest.param(
            {"Data": {"A1": ("n", "30")}},
            [("Data", ods_row(ods_cell("blank") + ods_cell("number", "30")))],
            {"Data": [[("blank", None), ("number", Decimal(30))]]},
            id="F6-coordinate-missing-despite-same-value-later",
        ),
        pytest.param(
            {"Data": {"A1": ("str", "")}},
            [("Data", ods_row(""))],
            {"Data": [[]]},
            id="F4-known-empty-missing-cell-is-not-empty",
        ),
        pytest.param(
            {"Data": {"A1": ("str", "")}},
            [("Data", ods_row(ods_cell("covered")))],
            {"Data": [[("covered", None)]]},
            id="F4-known-empty-covered-cell-is-not-empty",
        ),
        pytest.param(
            {"Data": {"A1": ("str", "")}},
            [("Data", ods_row(ods_cell("number", "0")))],
            {"Data": [[("number", Decimal(0))]]},
            id="F4-known-empty-number-zero-is-not-empty",
        ),
        pytest.param(
            {"Data": {"A1": ("str", "")}},
            [("Data", ods_row('<table:table-cell office:value="30"/>'))],
            {"Data": [[("untyped-scalar", "30")]]},
            id="F4-conflicting-untyped-scalar-is-not-blank",
        ),
        pytest.param(
            {"Data": {"A1": ("n", "30"), "A2": ("n", "30")}},
            [("Data", ods_row(ods_cell("number", "30", repeat=2)))],
            {"Data": [[("number", Decimal(30)), ("number", Decimal(30))]]},
            id="F6-column-repeats-do-not-satisfy-missing-next-row",
        ),
    ],
)
def test_misbound_or_mistyped_external_cache_never_reaches_final_publication(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    round_trip_runtime: Any,
    cells: SourceCells,
    sheets: list[tuple[str, str]],
    expected: TypedSheets,
) -> None:
    source = tmp_path / "linked.xlsx"
    original = _source(source, cells)
    delivery = tmp_path / "bad-backend-delivery.ods"
    write_ods(delivery, sheets)
    assert ods_typed_matrix(delivery) == expected
    prepared = tmp_path / "observed-private.xlsx"
    _backend(monkeypatch, delivery, cells, prepared)
    output = tmp_path / "published"
    output.mkdir()
    sentinel = output / "previous.txt"
    sentinel.write_bytes(b"previous successful result")
    result = round_trip_runtime.execute(_request(source, output, "typed-ods-negative"))
    assert not result.success
    assert result.error is not None
    assert "EXTERNAL_LINK_FLATTENING_NOT_DELIVERED" in {item.code for item in result.diagnostics}
    assert result.artifacts == []
    assert sorted(item.name for item in output.iterdir()) == [sentinel.name]
    assert sentinel.read_bytes() == b"previous successful result"
    assert_prepared_xlsx(prepared, cells)
    assert source.read_bytes() == original


@pytest.mark.parametrize(("kind", "value"), [("str", None), ("n", ""), ("n", None)])
def test_missing_cache_remains_distinct_from_known_empty_and_is_not_published(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, round_trip_runtime: Any, kind: str, value: str | None
) -> None:
    source = tmp_path / "missing-cache.xlsx"
    cells = {"Data": {"A1": (kind, value)}}
    original = _source(source, cells)
    delivery = tmp_path / "backend-delivery.ods"
    write_ods(delivery, [("Data", ods_row(ods_cell("string", "")))])
    prepared = tmp_path / "observed-private.xlsx"
    _backend(monkeypatch, delivery, cells, prepared)
    output = tmp_path / "published"
    result = round_trip_runtime.execute(_request(source, output, "typed-ods-missing-cache"))
    assert not result.success
    assert result.error is not None
    assert "EXTERNAL_LINK_CACHED_VALUE_MISSING" in {item.code for item in result.diagnostics}
    assert result.artifacts == []
    assert not output.exists() or list(output.iterdir()) == []
    assert not prepared.exists()
    assert source.read_bytes() == original
