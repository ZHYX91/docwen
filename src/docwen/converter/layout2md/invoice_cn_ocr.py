from __future__ import annotations

from typing import Any

from docwen.utils import ocr_utils

from . import invoice_cn


def parse_invoice_from_image(
    image_path: str,
    cancel_event: Any | None = None,
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    ocr_text = ocr_utils.extract_text_simple(image_path, cancel_event=cancel_event)

    compact = invoice_cn.compact_text(ocr_text)
    metadata = invoice_cn.parse_invoice_metadata_from_compact_text(compact)
    rows = invoice_cn.parse_invoice_rows_from_pdf_text(ocr_text, prefer_marked=True)
    if not rows:
        rows = invoice_cn.parse_invoice_rows_from_pdf_text(ocr_text, prefer_marked=False)

    return metadata, rows


def build_invoice_md_text(
    *,
    file_stem: str,
    metadata: dict[str, str | None],
    rows: list[dict[str, str]],
    include_empty: bool = True,
) -> str:
    headers = ["商品名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]

    metadata_yaml: dict[str, str | None] = {}
    for k in invoice_cn.INVOICE_CN_YAML_SCHEMA:
        v = metadata.get(k)
        metadata_yaml[k] = (str(v).strip() if v is not None else "")

    yaml_frontmatter = invoice_cn.build_yaml_frontmatter(
        file_stem=file_stem,
        metadata=metadata_yaml,
        include_empty=include_empty,
    )
    table_md = invoice_cn.render_markdown_table(headers=headers, rows=rows)
    return yaml_frontmatter + "## 商品明细\n\n" + table_md
