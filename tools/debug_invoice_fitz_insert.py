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

    from docwen_plugin_optimizer_invoice_cn.invoice_cn.metadata import (
        parse_invoice_metadata_from_compact_text,
    )
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import _compact_text

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
            compact = _compact_text(text)
            meta = parse_invoice_metadata_from_compact_text(compact)
            print(repr(text))
            print(repr(compact))
            print(meta)
        finally:
            d.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
