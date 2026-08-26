from __future__ import annotations

import sys
import tempfile
from pathlib import Path


def _ensure_workspace_packages_on_path() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    package_srcs = [
        repo_root / "packages" / "plugins" / "optimizers" / "invoice_cn" / "src",
        repo_root / "packages" / "core" / "src",
    ]
    for src in reversed(package_srcs):
        src_text = str(src)
        if src_text not in sys.path:
            sys.path.insert(0, src_text)


def main() -> int:
    _ensure_workspace_packages_on_path()

    import fitz

    from docwen_plugin_optimizer_invoice_cn.invoice_cn.markdown_writer import (
        _build_yaml_frontmatter,
        _render_markdown_table,
    )
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import parse_pdf_invoice
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import (
        INVOICE_CN_YAML_SCHEMA,
        TABLE_HEADERS,
    )

    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        pdf_path = tmp / "input.pdf"
        doc = fitz.open()
        p1 = doc.new_page()
        p1.insert_text((72, 72), "发票号码：11111111\n开票日期：2026年01月01日\n价税合计：1.00")
        p2 = doc.new_page()
        p2.insert_text((72, 72), "发票号码：22222222\n开票日期：2026年01月02日\n价税合计：2.00")
        doc.save(str(pdf_path))
        doc.close()

        out_dir = tmp / "out"
        out_dir.mkdir()

        metadata, rows = parse_pdf_invoice(str(pdf_path))
        metadata_yaml = {k: str(metadata.get(k) or "").strip() for k in INVOICE_CN_YAML_SCHEMA}
        md_text = (
            _build_yaml_frontmatter(
                file_stem="原始文件名",
                metadata=metadata_yaml,
                include_empty=True,
            )
            + "## 商品明细\n\n"
            + _render_markdown_table(headers=TABLE_HEADERS, rows=rows)
        )
        out_path = out_dir / "b_20260101_000000_fromPdf.md"
        out_path.write_text(md_text, encoding="utf-8")

        text = out_path.read_text(encoding="utf-8")
        print(out_path.name)
        print(text.splitlines()[0:20])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
