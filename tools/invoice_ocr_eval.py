from __future__ import annotations

import argparse
import json
import sys
import tempfile
from dataclasses import asdict, dataclass
from pathlib import Path


def _ensure_workspace_packages_on_path() -> None:
    try:
        import docwen_plugin_optimizer_invoice_cn  # noqa: F401

        return
    except Exception:
        pass

    repo_root = Path(__file__).resolve().parents[1]
    package_srcs = [
        repo_root / "packages" / "plugins" / "optimizers" / "invoice_cn" / "src",
        repo_root / "packages" / "core" / "src",
    ]
    for src in reversed(package_srcs):
        src_text = str(src)
        if src_text not in sys.path:
            sys.path.insert(0, src_text)


@dataclass(slots=True)
class FileReport:
    file: str
    pages: int
    pdf_text_len: int
    ocr_text_len: int | None
    mismatched_fields: list[str]
    text_fields: dict[str, str]
    ocr_fields: dict[str, str] | None
    notes: list[str]


def _normalize(s: str | None) -> str:
    return (s or "").strip()


def _mask_value(value: str) -> str:
    s = _normalize(value)
    if not s:
        return ""
    if len(s) <= 6:
        return "*" * len(s)
    return s[:2] + "*" * (len(s) - 4) + s[-2:]


def _safe_fields(fields: dict[str, str], *, unsafe_show_values: bool) -> dict[str, str]:
    if unsafe_show_values:
        return {k: _normalize(v) for k, v in fields.items()}
    return {k: _mask_value(v) for k, v in fields.items()}


def _render_first_page_to_png(pdf_path: str, png_path: str, *, dpi: int) -> None:
    import fitz

    doc = fitz.open(pdf_path)
    try:
        if doc.page_count < 1:
            raise RuntimeError("empty_pdf")
        page = doc.load_page(0)
        zoom = max(72, int(dpi)) / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        pix.save(png_path)
    finally:
        doc.close()


def _extract_from_pdf_text(pdf_path: str) -> tuple[dict[str, str], int, int]:
    import fitz

    from docwen_plugin_optimizer_invoice_cn.invoice_cn.pdf_parser import (
        parse_pdf_invoice,
        read_pdf_text_and_spans,
    )
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import (
        INVOICE_CN_YAML_SCHEMA,
    )

    doc = fitz.open(pdf_path)
    try:
        joined, _spans = read_pdf_text_and_spans(pdf_path)
        metadata, rows = parse_pdf_invoice(pdf_path)
        out: dict[str, str] = {}
        for k in INVOICE_CN_YAML_SCHEMA:
            out[k] = _normalize(str(metadata.get(k) or ""))
        out["_rows"] = str(len(rows))
        return out, doc.page_count, len(_normalize(joined))
    finally:
        doc.close()


def _extract_from_ocr_image(image_path: str) -> tuple[dict[str, str], int]:
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.image_parser import (
        parse_image_invoice,
        read_image_invoice_ocr_text,
    )
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import (
        INVOICE_CN_YAML_SCHEMA,
    )

    ocr_text = read_image_invoice_ocr_text(image_path)
    metadata, rows = parse_image_invoice(image_path)
    out: dict[str, str] = {}
    for k in INVOICE_CN_YAML_SCHEMA:
        out[k] = _normalize(str(metadata.get(k) or ""))
    out["_rows"] = str(len(rows))
    return out, len(_normalize(ocr_text))


def _compare_fields(a: dict[str, str], b: dict[str, str]) -> list[str]:
    keys = set(a.keys()) | set(b.keys())
    keys.discard("_rows")
    mismatched: list[str] = []
    for k in sorted(keys):
        if _normalize(a.get(k)) != _normalize(b.get(k)):
            mismatched.append(k)
    if _normalize(a.get("_rows")) != _normalize(b.get("_rows")):
        mismatched.append("_rows")
    return mismatched


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    parser.add_argument("--glob", default="**/*.pdf")
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--report", default="invoice_ocr_report.json")
    parser.add_argument("--unsafe-show-values", action="store_true")
    args = parser.parse_args()

    _ensure_workspace_packages_on_path()

    root = Path(args.path)
    files = [p for p in sorted(root.glob(args.glob)) if p.is_file()]
    if args.limit and args.limit > 0:
        files = files[: int(args.limit)]

    reports: list[FileReport] = []
    for f in files:
        text_fields, pages, pdf_text_len = _extract_from_pdf_text(str(f))
        notes: list[str] = []
        ocr_fields: dict[str, str] | None = None
        ocr_text_len: int | None = None
        mismatched: list[str] = []

        try:
            with tempfile.TemporaryDirectory() as td:
                png_path = str(Path(td) / "page_1.png")
                _render_first_page_to_png(str(f), png_path, dpi=int(args.dpi))
                ocr_fields, ocr_text_len = _extract_from_ocr_image(png_path)
                mismatched = _compare_fields(text_fields, ocr_fields)
        except (ModuleNotFoundError, ImportError) as e:
            name = getattr(e, "name", None)
            notes.append(f"missing_dependency:{name or type(e).__name__}")
        except Exception as e:
            notes.append(f"ocr_failed:{type(e).__name__}")

        reports.append(
            FileReport(
                file=f.name,
                pages=pages,
                pdf_text_len=pdf_text_len,
                ocr_text_len=ocr_text_len,
                mismatched_fields=mismatched,
                text_fields=_safe_fields(text_fields, unsafe_show_values=bool(args.unsafe_show_values)),
                ocr_fields=(
                    _safe_fields(ocr_fields, unsafe_show_values=bool(args.unsafe_show_values)) if ocr_fields else None
                ),
                notes=notes,
            )
        )

    out = {
        "schema_version": 1,
        "root": str(root),
        "glob": str(args.glob),
        "dpi": int(args.dpi),
        "count": len(reports),
        "reports": [asdict(r) for r in reports],
    }
    report_path = Path(args.report)
    if not report_path.is_absolute():
        report_path = root / report_path
    report_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(report_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
