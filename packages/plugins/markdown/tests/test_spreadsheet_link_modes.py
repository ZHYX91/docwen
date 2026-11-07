"""Native single-cell links, fallback losslessness and YAML parity."""

from __future__ import annotations

import csv
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from docwen_plugin_markdown.to_spreadsheet.converter import MdToCsvConverter, MdToXlsxConverter

from .conftest import make_context

pytestmark = pytest.mark.contract


def active(workbook: Workbook) -> Worksheet:
    sheet = workbook.active
    assert isinstance(sheet, Worksheet)
    return sheet


def convert(
    tmp_path: Path, content: str, *, target: str = "xlsx", config: dict | None = None, template: Path | None = None
):
    source = tmp_path / "links.md"
    source.write_text(content, encoding="utf-8")
    context, _ = make_context(
        str(source),
        target_format=target,
        config_values={"link": {"non_embed_links": config or {}}},
        options={"template_name": str(template)} if template else {},
    )
    result = (MdToXlsxConverter() if target == "xlsx" else MdToCsvConverter()).convert(context)
    assert result.success, result.error
    return result.artifacts[0].staging_path


@pytest.mark.parametrize("syntax", ["[Site](https://example.com)", "[[https://example.com|Site]]"])
def test_single_link_is_real_hyperlink_after_reload(tmp_path: Path, syntax: str) -> None:
    path = convert(tmp_path, "| Link |\n| --- |\n| " + syntax.replace("|", r"\|") + " |\n")
    cell = active(load_workbook(path))["A2"]
    assert cell.value == "Site"
    assert cell.hyperlink.target == "https://example.com"


@pytest.mark.parametrize(
    "fallback,expected",
    [
        ("keep", "before [Site](https://example.com) after"),
        ("extract_text", "before Site after"),
        ("extract_address", "before https://example.com after"),
    ],
)
def test_mixed_link_fallback_does_not_split_cell(tmp_path: Path, fallback: str, expected: str) -> None:
    path = convert(
        tmp_path,
        "| Link |\n| --- |\n| before [Site](https://example.com) after |\n",
        config={"markdown_hyperlink_fallback": fallback},
    )
    sheet = active(load_workbook(path))
    assert sheet.max_column == 1 and sheet.max_row == 2
    assert sheet["A2"].value == expected
    assert sheet["A2"].hyperlink is None


def test_multiple_links_preserve_both_targets(tmp_path: Path) -> None:
    value = "[A](https://a.example) [B](https://b.example)"
    path = convert(tmp_path, f"| Links |\n| --- |\n| {value} |\n")
    cell = active(load_workbook(path))["A2"]
    assert cell.value == value and cell.hyperlink is None


@pytest.mark.parametrize(
    "mode,expected",
    [
        ("keep", "[x](<https://example.com/a>)"),
        ("extract_text", "x"),
        ("extract_address", "https://example.com/a"),
        ("remove", None),
        ("hyperlink", "x"),
    ],
)
def test_yaml_five_modes_and_template_literal_is_not_reparsed(tmp_path: Path, mode: str, expected: str | None) -> None:
    template = tmp_path / "template.xlsx"
    wb = Workbook()
    active(wb)["A1"] = "{{link}}"
    active(wb)["A2"] = "[template](https://template.example)"
    wb.save(template)
    path = convert(
        tmp_path, "---\nlink: '[x](<https://example.com/a>)'\n---\n", config={"markdown_mode": mode}, template=template
    )
    sheet = active(load_workbook(path))
    assert sheet["A1"].value == expected
    assert bool(sheet["A1"].hyperlink) == (mode == "hyperlink")
    assert sheet["A2"].value == "[template](https://template.example)"
    assert sheet["A2"].hyperlink is None


@pytest.mark.parametrize(
    "fallback,expected",
    [("keep", "[Site](https://example.com)"), ("extract_text", "Site"), ("extract_address", "https://example.com")],
)
def test_csv_uses_selected_fallback(tmp_path: Path, fallback: str, expected: str) -> None:
    path = convert(
        tmp_path,
        "| Link |\n| --- |\n| [Site](https://example.com) |\n",
        target="csv",
        config={"markdown_hyperlink_fallback": fallback},
    )
    with open(path, encoding="utf-8-sig", newline="") as handle:
        assert list(csv.reader(handle)) == [["Link"], [expected]]


def test_extract_address_does_not_make_formula_or_resolve_private_path(tmp_path: Path) -> None:
    path = convert(tmp_path, "| Link |\n| --- |\n| [x](=1+1) |\n", config={"markdown_mode": "extract_address"})
    cell = active(load_workbook(path))["A2"]
    assert cell.value == "=1+1" and cell.data_type == "s"


def test_yaml_list_multiple_links_uses_fallback(tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    wb = Workbook()
    active(wb)["A1"] = "{{links}}"
    wb.save(template)
    path = convert(
        tmp_path,
        "---\nlinks:\n  - '[A](https://a.example)'\n  - '[B](https://b.example)'\n---\n",
        config={"markdown_hyperlink_fallback": "extract_address"},
        template=template,
    )
    cell = active(load_workbook(path))["A1"]
    assert cell.value == "https://a.example, https://b.example" and cell.hyperlink is None


def test_private_link_token_does_not_reduce_cell_capacity(tmp_path: Path) -> None:
    value = "x" * (32767 - len("[x](https://e.co)")) + "[x](https://e.co)"
    path = convert(tmp_path, f"| Links |\n| --- |\n| {value} |\n")
    cell = active(load_workbook(path))["A2"]
    assert cell.value == value and cell.hyperlink is None


def test_hyperlink_label_formula_is_literal(tmp_path: Path) -> None:
    path = convert(tmp_path, "| Links |\n| --- |\n| [=1+1](https://example.com) |\n")
    cell = active(load_workbook(path))["A2"]
    assert cell.value == "=1+1" and cell.data_type == "s"
    assert cell.hyperlink.target == "https://example.com"


def test_link_in_header_retains_template_column_lookup(tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    wb = Workbook()
    active(wb)["A1"] = "{{↓[Site](https://example.com)}}"
    wb.save(template)
    path = convert(tmp_path, "| [Site](https://example.com) |\n| --- |\n| preserved |\n", template=template)
    assert active(load_workbook(path))["A1"].value == "preserved"


@pytest.mark.parametrize("syntax", ["[[target.md|Local]]", "[Local](target.md)"])
def test_embedded_links_retain_child_source_base(tmp_path: Path, syntax: str) -> None:
    folder = tmp_path / "nested"
    folder.mkdir()
    target = folder / "target.md"
    target.write_text("destination", encoding="utf-8")
    (folder / "child.md").write_text(f"| Link |\n| --- |\n| {syntax} |\n", encoding="utf-8")
    path = convert(tmp_path, "![[nested/child.md]]")
    cell = active(load_workbook(path))["A2"]
    assert cell.value == "Local"
    assert cell.hyperlink.target == target.as_uri()


def test_empty_label_uses_authored_target_not_private_absolute_path(tmp_path: Path) -> None:
    (tmp_path / "target.md").write_text("target", encoding="utf-8")
    path = convert(tmp_path, "| Links |\n| --- |\n| [](target.md) |\n")
    cell = active(load_workbook(path))["A2"]
    assert cell.value == "target.md"
    assert cell.hyperlink.target == (tmp_path / "target.md").as_uri()


@pytest.mark.parametrize("target", [r"javascript\:alert(1)", r"data\:text/plain,hi", "javascript:alert(1)"])
def test_dangerous_decoded_scheme_never_becomes_native_hyperlink(tmp_path: Path, target: str) -> None:
    syntax = f"[Label]({target})"
    path = convert(tmp_path, f"| Links |\n| --- |\n| {syntax} |\n")
    cell = active(load_workbook(path))["A2"]
    assert cell.value == syntax
    assert cell.hyperlink is None


@pytest.mark.parametrize("embedded", ["![[missing.md]]", "![[missing.png]]", "![x](missing.png)"])
def test_yaml_embeds_are_never_looked_up_or_rewritten(
    tmp_path: Path, embedded: str, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_core.export_semantics import LinkRuntimeConfig
    from docwen_core.links import _non_embed
    from docwen_plugin_markdown.to_spreadsheet.links import SpreadsheetLinks

    def forbidden(*args, **kwargs):
        raise AssertionError("YAML embed lookup is forbidden")

    monkeypatch.setattr(_non_embed, "resolve_file_path", forbidden)
    links = SpreadsheetLinks(str(tmp_path / "source.md"), LinkRuntimeConfig(), native=True)
    assert links.project(embedded) == embedded


def test_wiki_alias_punctuation_remains_literal_in_native_link(tmp_path: Path) -> None:
    path = convert(tmp_path, "| Link |\n| --- |\n| [[https://example.com\\|*literal*]] |\n")
    cell = active(load_workbook(path))["A2"]
    assert cell.value == "*literal*"
    assert cell.hyperlink.target == "https://example.com"


@pytest.mark.parametrize("target", ["xlsx", "csv"])
@pytest.mark.parametrize("mode", ["extract_text", "hyperlink"])
def test_formatted_markdown_label_is_plain_in_text_projection(tmp_path: Path, target: str, mode: str) -> None:
    path = convert(
        tmp_path,
        "| Link |\n| --- |\n| prefix [**Bold** `code`](https://example.com) suffix |\n",
        target=target,
        config={"markdown_mode": mode, "markdown_hyperlink_fallback": "extract_text"},
    )
    if target == "xlsx":
        value = active(load_workbook(path))["A2"].value
    else:
        with open(path, encoding="utf-8-sig", newline="") as handle:
            value = list(csv.reader(handle))[1][0]
    assert value == "prefix Bold code suffix"


@pytest.mark.parametrize("target", ["xlsx", "csv"])
@pytest.mark.parametrize("mode", ["extract_text", "hyperlink"])
@pytest.mark.parametrize("label", ["Energy $mc^2$ result", "$mc^2$"])
def test_formula_link_label_survives_spreadsheet_projection(tmp_path: Path, target: str, mode: str, label: str) -> None:
    path = convert(
        tmp_path,
        f"| Link |\n| --- |\n| [{label}](https://example.com) |\n",
        target=target,
        config={"markdown_mode": mode, "markdown_hyperlink_fallback": "extract_text"},
    )
    if target == "xlsx":
        cell = active(load_workbook(path))["A2"]
        assert cell.value == label
        assert cell.data_type == "s"
        assert getattr(cell.hyperlink, "target", None) == ("https://example.com" if mode == "hyperlink" else None)
    else:
        with open(path, encoding="utf-8-sig", newline="") as handle:
            assert list(csv.reader(handle)) == [["Link"], [label]]
