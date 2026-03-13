from __future__ import annotations

import tempfile
from pathlib import Path


def main() -> int:
    import fitz

    from docwen.converter.layout2md import invoice_cn

    with tempfile.TemporaryDirectory() as td:
        pdf = Path(td) / "m.pdf"
        doc = fitz.open()
        p1 = doc.new_page()
        p1.insert_text((72, 72), "发票号码：11111111\n开票日期：2026年01月01日\n价税合计：1.00")
        doc.save(str(pdf))
        doc.close()

        d = fitz.open(str(pdf))
        try:
            page = d[0]
            text = str(page.get_text("text"))
            compact = invoice_cn.compact_text(text)
            meta = invoice_cn.parse_invoice_metadata_from_compact_text(compact)
            print(repr(text))
            print(repr(compact))
            print(meta)
        finally:
            d.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
