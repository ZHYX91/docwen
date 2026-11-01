"""New synthetic text PDFs retain signed fields, party roles and page identity."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

import fitz
import pytest
import yaml
from markdown_it import MarkdownIt

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_plugin_optimizer_invoice_cn.plugin import InvoicePlugin
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]
_BUYER = "123456789012345678"
_SELLER = "987654321098765432"
_HEADERS = ["商品名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]
_X = [40.0, 200.0, 270.0, 330.0, 430.0, 560.0, 690.0, 790.0]
_Item = tuple[float, float, str]


def _parties(mode: str, *, buyer: str = _BUYER, seller: str = _SELLER) -> list[_Item]:
    parties = []
    for role, name, identity, x, y in (
        ("购买方", "购买单位", buyer, 40.0, 110.0),
        (
            "销售方",
            "销售单位",
            seller,
            520.0 if "horizontal" in mode else 40.0,
            110.0 if "horizontal" in mode else 195.0,
        ),
    ):
        if (mode == "buyer-only" and role == "销售方") or (mode == "seller-only" and role == "购买方"):
            continue
        if mode != "unknown":
            parties.append((x, y, role + "信息"))
        parties.append((x, y + 18.0, "名称：" + name))
        if "split" in mode:
            parties.extend([(x, y + 36.0, "统一社会信用代码"), (x + 82.0, y + 36.0, "/纳税人识别号：")])
        else:
            parties.append((x, y + 36.0, "统一社会信用代码/纳税人识别号："))
        parties.append((x + 185.0, y + 36.0, identity))
    return sorted(parties, key=lambda item: (item[1], item[0])) if mode == "horizontal-interleaved" else parties


def _page_items(
    *,
    mode: str = "vertical",
    amount: str = "-100.00",
    tax: str = "-13.00",
    total: str = "-113.00",
    quantity: str = "-1",
    unit_price: str = "100.00",
    item_name: str = "服务甲",
    buyer: str = _BUYER,
    seller: str = _SELLER,
    include_totals: bool = True,
) -> tuple[list[_Item], list[str]]:
    items = [
        (40.0, 40.0, "电子发票（普通发票）"),
        (40.0, 58.0, "发票号码：87654321"),
        (40.0, 76.0, "开票日期：2024年01月15日"),
    ]
    items.extend(_parties(mode, buyer=buyer, seller=seller))
    items.extend(
        (x, 300.0, "项目名称" if header == "商品名称" else header) for x, header in zip(_X, _HEADERS, strict=True)
    )
    row = [item_name, "型号A", "件", quantity, unit_price, amount, "13%", tax]
    items.extend((x, 330.0, value) for x, value in zip(_X, row, strict=True))
    if include_totals:
        items.extend(
            [
                (40.0, 370.0, "合计"),
                (560.0, 370.0, "￥" + amount),
                (790.0, 370.0, "￥" + tax),
                (40.0, 400.0, "价税合计（小写）"),
                (560.0, 400.0, "￥" + total),
            ]
        )
    return items, row


def _pdf(path: Path, pages: list[list[_Item]]) -> None:
    with fitz.open() as document:
        for items in pages:
            page = document.new_page(width=1000, height=600)
            for x, y, text in items:
                page.insert_text((x, y), text, fontname="china-s", fontsize=8)
        document.save(str(path))
    with fitz.open(path) as document:
        assert len(document) == len(pages)
        for page, items in zip(document, pages, strict=True):
            native_lines = str(page.get_text("text")).splitlines()
            assert native_lines == [text for _x, _y, text in items]


def _run(source: Path, root: Path) -> tuple[dict[str, Any], list[list[str]], Any]:
    before = hashlib.sha256(source.read_bytes()).digest()
    output_dir = root / "out"
    output_dir.mkdir()
    registry = PluginRegistry()
    registry.register(InvoicePlugin())
    runtime = TaskManager(
        registry, RouteResolver(registry), WorkspaceManager(root_dir=str(root / "workspace")), OutputFinalizer()
    )
    result = runtime.execute_single(
        ConversionRequest(
            request_id="native-invoice-fields",
            input_refs=[
                FileRef(
                    path=str(source),
                    format=source.suffix.lstrip("."),
                    category="layout",
                    size_bytes=source.stat().st_size,
                )
            ],
            target_format="md",
            action_name="invoice_cn",
            options={"to_md_enable_ocr": False},
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )
    )
    assert result.success, result.error
    assert hashlib.sha256(source.read_bytes()).digest() == before
    [artifact] = [artifact for artifact in result.artifacts if artifact.media_type == "text/markdown"]
    output = Path(artifact.staging_path)
    assert output.is_file() and output.is_relative_to(output_dir)
    assert artifact.metadata["document_node_committed"] is True
    assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
    assert not any(d.code in {"INVOICE-OCR-OK", "INVOICE-SCAN-DETECTED"} for d in result.diagnostics)
    _, header, body = output.read_text(encoding="utf-8").split("---", 2)
    metadata = yaml.safe_load(header)
    tokens = MarkdownIt("commonmark").enable("table").parse(body)
    rows: list[list[str]] = []
    current: list[str] | None = None
    for token in tokens:
        if token.type == "tr_open":
            current = []
        elif token.type == "inline" and current is not None:
            current.append("".join(child.content for child in (token.children or [])))
        elif token.type == "tr_close":
            assert current is not None
            rows.append(current)
            current = None
    assert rows[0] == _HEADERS
    assert all(len(row) == 8 for row in rows)
    return metadata, rows[1:], result


@pytest.mark.parametrize(
    "amount,tax,total,quantity,unit_price",
    [
        ("-100.00", "-13.00", "-113.00", "-1", "100.00"),
        ("100.00", "13.00", "113.00", "1", "100.00"),
        ("1.00", "9.00", "10.00", "1", "1.00"),
        ("-5.00", "-20.00", "-25.00", "-1", "5.00"),
        ("9007199254740993.01", "9007199254740993.02", "18014398509481986.03", "1", "9007199254740993.01"),
        ("0.00", "0.00", "0.00", "0", "100.00"),
        ("0.12345", "0.00001", "0.12346", "1", "0.12345"),
    ],
)
def test_signed_amount_fields_follow_roles_not_magnitude_or_positive_unit_price(
    tmp_path: Path, amount: str, tax: str, total: str, quantity: str, unit_price: str
) -> None:
    items, row = _page_items(amount=amount, tax=tax, total=total, quantity=quantity, unit_price=unit_price)
    source = tmp_path / "signed-fields.pdf"
    _pdf(source, [items])
    metadata, rows, _ = _run(source, tmp_path)
    assert {field: metadata[field] for field in ("金额", "税额", "价税合计")} == {
        "金额": amount,
        "税额": tax,
        "价税合计": total,
    }
    assert all(isinstance(metadata[field], str) for field in ("金额", "税额", "价税合计"))
    assert rows == [row]


@pytest.mark.parametrize(
    "mode",
    [
        "vertical",
        "horizontal",
        "horizontal-interleaved",
        "vertical-split",
        "horizontal-split",
        "buyer-only",
        "seller-only",
        "unknown",
    ],
)
def test_party_ids_have_one_proven_role_in_the_published_yaml(tmp_path: Path, mode: str) -> None:
    items, row = _page_items(mode=mode)
    source = tmp_path / "party-fields.pdf"
    _pdf(source, [items])
    metadata, rows, result = _run(source, tmp_path)
    expected_buyer = "" if mode in {"seller-only", "unknown"} else _BUYER
    expected_seller = "" if mode in {"buyer-only", "unknown"} else _SELLER
    assert metadata["购买方纳税人识别号"] == expected_buyer
    assert metadata["销售方纳税人识别号"] == expected_seller
    assert rows == [row]
    if mode == "unknown":
        assert any(
            d.code == "INVOICE-PARTY-AMBIGUOUS" and d.location == "party-fields.pdf:page-1" for d in result.diagnostics
        )


def test_same_coordinates_on_different_pages_never_share_party_or_detail_rows(tmp_path: Path) -> None:
    first, first_row = _page_items(
        mode="buyer-only",
        amount="-40.00",
        tax="-5.20",
        unit_price="40.00",
        item_name="第一页服务",
        include_totals=False,
    )
    second, second_row = _page_items(mode="seller-only", item_name="第二页服务")
    # Keep the visible seller role and value at exactly the first page's buyer coordinates.
    second = [(x, y - 85.0 if 195.0 <= y <= 231.0 else y, text) for x, y, text in second]
    source = tmp_path / "page-identities.pdf"
    _pdf(source, [first, second])
    metadata, rows, result = _run(source, tmp_path)
    assert (metadata["购买方纳税人识别号"], metadata["销售方纳税人识别号"]) == (_BUYER, _SELLER)
    assert rows == [first_row, second_row]
    assert {field: metadata[field] for field in ("金额", "税额", "价税合计")} == {
        "金额": "-100.00",
        "税额": "-13.00",
        "价税合计": "-113.00",
    }
    assert not any(d.code == "INVOICE-METADATA-CONFLICT" for d in result.diagnostics)


def test_conflicting_page_metadata_keeps_first_owned_value_and_reports_both_pages(tmp_path: Path) -> None:
    first, first_row = _page_items(item_name="第一页服务")
    second, second_row = _page_items(item_name="第二页服务", buyer="111111111111111111")
    source = tmp_path / "page-conflict.pdf"
    _pdf(source, [first, second])
    metadata, rows, result = _run(source, tmp_path)
    assert (metadata["购买方纳税人识别号"], metadata["销售方纳税人识别号"]) == (_BUYER, _SELLER)
    assert rows == [first_row, second_row]
    [diagnostic] = [d for d in result.diagnostics if d.code == "INVOICE-METADATA-CONFLICT"]
    assert diagnostic.location == "page-conflict.pdf:page-2:购买方纳税人识别号"
    assert "page-1" in diagnostic.message and "page-2" in diagnostic.message


def test_complete_party_text_is_not_replaced_by_conflicting_geometry(tmp_path: Path) -> None:
    items, row = _page_items(mode="horizontal")
    # Keep complete role-labelled text order, but deliberately place the two
    # ID values in the opposite visual columns. This is conflicting evidence,
    # not permission for the position heuristic to silently replace text.
    items = [(705.0 if text == _BUYER else 225.0 if text == _SELLER else x, y, text) for x, y, text in items]
    source = tmp_path / "party-evidence-conflict.pdf"
    _pdf(source, [items])
    metadata, rows, result = _run(source, tmp_path)
    assert (metadata["购买方纳税人识别号"], metadata["销售方纳税人识别号"]) == (_BUYER, _SELLER)
    assert rows == [row]
    conflicts = [diagnostic for diagnostic in result.diagnostics if diagnostic.code == "INVOICE-METADATA-CONFLICT"]
    assert sorted(diagnostic.location for diagnostic in conflicts) == [
        "party-evidence-conflict.pdf:page-1:购买方纳税人识别号",
        "party-evidence-conflict.pdf:page-1:销售方纳税人识别号",
    ]


def test_complete_party_names_without_colons_retain_the_existing_text_contract(tmp_path: Path) -> None:
    items, row = _page_items(mode="vertical")
    items = [(x, y, text.replace("名称：", "名称")) for x, y, text in items]
    source = tmp_path / "party-labels-without-colons.pdf"
    _pdf(source, [items])
    metadata, rows, result = _run(source, tmp_path)
    assert (metadata["购买方名称"], metadata["销售方名称"]) == ("购买单位", "销售单位")
    assert (metadata["购买方纳税人识别号"], metadata["销售方纳税人识别号"]) == (_BUYER, _SELLER)
    assert rows == [row]
    assert not any(d.code in {"INVOICE-METADATA-CONFLICT", "INVOICE-PARTY-AMBIGUOUS"} for d in result.diagnostics)
