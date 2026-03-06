"""converter 单元测试。"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docwen.converter.layout2md.invoice_cn import convert_invoice_cn_layout_to_md

pytestmark = pytest.mark.unit


def test_invoice_cn_ofd_fallback_parses_text_from_content_xml(tmp_path: Path) -> None:
    ofd_path = tmp_path / "inv.ofd"

    nodes: list[str] = ['<?xml version="1.0" encoding="UTF-8"?>', "<Content>"]

    def add(x: float, y: float, text: str) -> None:
        nodes.append(f'  <TextObject Boundary="{x} {y} 1 1"><TextCode X="0" Y="0">{text}</TextCode></TextObject>')

    add(0, 10, "发票号码：99990000123456789012")
    add(0, 20, "开票日期：2026年02月11日")
    add(4.5, 35.0, "购买方信息")
    add(100.0, 35.0, "销售方信息")
    add(11.0, 39.67, "名称：")
    add(20.0, 39.42, "张三")
    add(111.5, 39.67, "名称：")
    add(120.0, 39.42, "示例商贸有限公司")
    add(4.5, 44.0, "统一社会信用代码/纳税人识别号：")
    add(50.0, 44.0, "911111111111111111")
    add(100.0, 44.0, "统一社会信用代码/纳税人识别号：")
    add(150.0, 44.0, "922222222222222222")

    add(16.65, 55.67, "项目名称")
    add(42.00, 55.67, "规格型号")
    add(66.74, 55.67, "单 位")
    add(92.97, 55.67, "数 量")
    add(117.97, 55.67, "单 价")
    add(143.97, 55.67, "金 额")
    add(157.27, 55.67, "税率/征收率")
    add(195.97, 55.67, "税 额")

    add(4.50, 58.92, "*商品类A*示例商品甲")
    add(41.50, 58.92, "1×0.8kg")
    add(70.00, 58.92, "盒")
    add(101.00, 58.92, "1")
    add(120.00, 58.92, "104.6")
    add(144.50, 58.92, "104.60")
    add(163.00, 58.92, "免税")
    add(201.00, 58.92, "***")
    add(4.50, 62.67, "续行A")

    add(4.50, 66.42, "*商品类B*示例商品乙")
    add(41.50, 66.42, "1×1盒")
    add(70.00, 66.42, "盒")
    add(101.00, 66.42, "1")
    add(118.50, 66.42, "106.11")
    add(144.50, 66.42, "106.11")
    add(163.75, 66.42, "13%")
    add(198.00, 66.42, "13.79")
    add(4.50, 70.16, "续行B1")
    add(4.50, 73.91, "续行B2")

    add(4.50, 77.66, "*商品类C*示例商品丙 1.5kg")
    add(41.50, 77.66, "1×1盒")
    add(70.00, 77.66, "盒")
    add(101.00, 77.66, "1")
    add(121.50, 77.66, "29.9")
    add(146.00, 77.66, "29.90")
    add(163.00, 77.66, "免税")
    add(201.00, 77.66, "***")

    add(4.50, 81.4, "*服务类*配送服务费")
    add(101.00, 81.4, "1")
    add(121.50, 81.4, "1.77")
    add(147.50, 81.4, "1.77")
    add(163.75, 81.4, "13%")
    add(199.50, 81.4, "0.23")

    add(0, 90, "合    计¥242.38¥14.02")
    add(0, 92, "价税合计（大写）贰佰伍拾陆圆肆角整（小写）¥256.40")

    nodes.append("</Content>")
    content_xml = "\n".join(nodes)

    with zipfile.ZipFile(ofd_path, "w") as zf:
        zf.writestr("OFD.xml", "<OFD></OFD>")
        zf.writestr("Doc_0/Pages/Page_0/Content.xml", content_xml.encode("utf-8"))

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    result = convert_invoice_cn_layout_to_md(
        file_path=str(ofd_path),
        actual_format="ofd",
        output_dir=str(out_dir),
        basename_for_output="inv",
        original_file_stem="99990000123456789012",
    )

    md_text = Path(result["md_path"]).read_text(encoding="utf-8")
    assert "发票号码: '99990000123456789012'" in md_text
    assert "优化类型:" not in md_text
    assert "购买方名称: 张三" in md_text
    assert "销售方名称: 示例商贸有限公司" in md_text
    assert "购买方纳税人识别号: '911111111111111111'" in md_text
    assert "销售方纳税人识别号: '922222222222222222'" in md_text
    assert "购买方地址电话: ''" in md_text
    assert "销售方开户行及账号: ''" in md_text
    assert "价税合计: '256.40'" in md_text
    assert "金额: '242.38'" in md_text
    assert "税额: '14.02'" in md_text
    assert "| *商品类A*示例商品甲续行A |" in md_text
    assert "| *商品类B*示例商品乙续行B1续行B2 |" in md_text
    assert "| *商品类C*示例商品丙 1.5kg |" in md_text
    assert "| *服务类*配送服务费 |" in md_text
