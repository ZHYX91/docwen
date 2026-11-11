"""Template CSV publication uses bounded ordinary-sheet plans through Runtime."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_runtime.templates import TemplateRegistry
from docwen_runtime.templates.state import user_templates_dir

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize("outcome", ["success", "limit", "cancel"])
def test_real_template_csv_publication_and_failure_cleanup(
    tmp_path: Path, round_trip_runtime: Any, monkeypatch: pytest.MonkeyPatch, outcome: str
) -> None:
    from docwen_plugin_markdown.to_spreadsheet import converter, template_xlsx

    directory = user_templates_dir()
    directory.mkdir(parents=True, exist_ok=True)
    template = directory / "bounded-csv.xlsx"
    workbook = Workbook()
    first = workbook.active
    assert first is not None
    first.title = "First"
    first["A1"] = "{{name}}"
    first["B2"] = "first last"
    second = workbook.create_sheet("Second")
    second.cell(1, 3001, "second last") if outcome == "cancel" else second.cell(2, 2, "second last")
    chart = BarChart()
    chart.add_data(Reference(first, min_col=2, min_row=1, max_row=2))
    workbook.create_chartsheet("Plot", 0).add_chart(chart)
    workbook.save(template)
    original_template = template.read_bytes()
    template_id = next(item.id for item in TemplateRegistry.default().list_templates("xlsx") if item.path == template)
    source = tmp_path / "source.md"
    source.write_text("---\nname: Filled\n---\n\n# Source\n", encoding="utf-8")
    original_source = source.read_bytes()
    output = tmp_path / "published"
    output.mkdir()
    sentinel = output / "existing.csv"
    sentinel.write_bytes(b"existing output")
    request_id = f"template-csv-{outcome}"
    if outcome == "limit":
        monkeypatch.setattr(converter, "MAX_SPREADSHEET_PROJECTION_CELLS", 6)
    visits = 0
    if outcome == "cancel":
        load = template_xlsx.load_template_workbook

        class Values(dict):
            def get(self, key, default=None):
                nonlocal visits
                visits += 1
                if visits == 17:
                    round_trip_runtime.cancel(request_id)
                return super().get(key, default)

        def capture(path):
            view = load(path)
            view["Second"]._cells = Values(view["Second"]._cells)
            return view

        monkeypatch.setattr(template_xlsx, "load_template_workbook", capture)

    result = round_trip_runtime.execute(
        ConversionRequest(
            request_id=request_id,
            input_refs=[FileRef(str(source), "markdown", "markdown", len(original_source))],
            target_format="csv",
            options={"template_name": template_id},
            output_policy=OutputPolicy(output_dir=str(output)),
        )
    )

    assert template.read_bytes() == original_template
    assert source.read_bytes() == original_source
    assert sentinel.read_bytes() == b"existing output"
    if outcome == "success":
        assert result.success, result.error
        assert [item.metadata["sheet_name"] for item in result.artifacts] == ["First", "Second"]
        assert [item.metadata["sheet_index"] for item in result.artifacts] == [0, 1]
        assert [item.metadata["sheet_count"] for item in result.artifacts] == [2, 2]
        assert [item.is_primary for item in result.artifacts] == [True, False]
        matrices = []
        for artifact in result.artifacts:
            final = Path(artifact.staging_path)
            assert final.is_relative_to(output)
            with final.open(encoding="utf-8-sig", newline="") as handle:
                matrices.append(list(csv.reader(handle)))
        assert matrices == [[["Filled", ""], ["", "first last"]], [["", ""], ["", "second last"]]]
    else:
        assert not result.success and result.error is not None and result.artifacts == []
        if outcome == "limit":
            assert result.error.diagnostic_code == "SPREADSHEET-PROJECTION-LIMIT"
        else:
            assert result.error.error_type == "cancelled"
            assert 17 <= visits <= 1000
        assert list(output.rglob("*")) == [sentinel]
