"""Canonical YAML frontmatter schema for Chinese invoices."""

from __future__ import annotations

INVOICE_CN_YAML_SCHEMA: list[str] = [
    "发票种类",
    "发票代码",
    "发票号码",
    "开票日期",
    "校验码",
    "购买方名称",
    "购买方纳税人识别号",
    "购买方地址电话",
    "购买方开户行及账号",
    "销售方名称",
    "销售方纳税人识别号",
    "销售方地址电话",
    "销售方开户行及账号",
    "金额",
    "税额",
    "价税合计",
    "备注",
    "收款人",
    "复核",
    "开票人",
]

TABLE_HEADERS: list[str] = [
    "商品名称",
    "规格型号",
    "单位",
    "数量",
    "单价",
    "金额",
    "税率",
    "税额",
]
