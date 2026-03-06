from __future__ import annotations

import tempfile
from pathlib import Path


def main() -> int:
    import fitz

    from docwen.converter.layout2md.invoice_cn import convert_invoice_cn_layout_to_md

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

        result = convert_invoice_cn_layout_to_md(
            file_path=str(pdf_path),
            actual_format="pdf",
            output_dir=str(out_dir),
            basename_for_output="b_20260101_000000_fromPdf",
            original_file_stem="原始文件名",
        )

        for p in result["md_paths"]:
            text = Path(p).read_text(encoding="utf-8")
            print(Path(p).name)
            print(text.splitlines()[0:20])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

