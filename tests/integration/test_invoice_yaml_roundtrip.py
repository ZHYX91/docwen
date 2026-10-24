"""Invoice output keeps identifiers and field strings through YAML consumption."""

from __future__ import annotations

import pytest

from docwen_plugin_markdown.yaml_processor import extract_yaml_front_matter
from docwen_plugin_optimizer_invoice_cn.invoice_cn.markdown_writer import _build_yaml_frontmatter

pytestmark = pytest.mark.integration


@pytest.mark.parametrize(
    "invoice_number",
    ["00000001", "00123456", "2604700000000392402353", "on", "off", "-1", "1e3"],
)
def test_invoice_fields_retain_string_values_and_types(invoice_number: str) -> None:
    metadata = {
        "发票号码": invoice_number,
        "开票日期": "2026-10-24",
        "价税合计": "1234.00",
        "备注": 'Path: C:\\new\\docs; it\'s "quoted"\nnext line',
    }
    content = _build_yaml_frontmatter(file_stem=invoice_number, metadata=metadata) + "\nInvoice body.\n"
    parsed, body = extract_yaml_front_matter(content)
    assert parsed == {"aliases": [invoice_number], "标题": invoice_number, **metadata}
    assert all(type(parsed[key]) is str for key in metadata)
    assert body.strip() == "Invoice body."


@pytest.mark.parametrize("include_empty", [False, True])
def test_invoice_empty_field_policy_and_localized_title_are_preserved(include_empty: bool) -> None:
    content = _build_yaml_frontmatter(
        file_stem="00123456",
        metadata={"missing": None, "blank": " \t", "value": " 00000001 "},
        include_empty=include_empty,
        yaml_key_labels={"title": "Titel"},
    )
    parsed, _body = extract_yaml_front_matter(content)
    expected = {"aliases": ["00123456"], "Titel": "00123456", "value": "00000001"}
    if include_empty:
        expected.update({"missing": "", "blank": ""})
    assert parsed == expected
