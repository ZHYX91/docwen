from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path


@dataclass(slots=True)
class CheckResult:
    file: str
    ok: bool
    missing_fields: int
    total_mismatch: bool


_FRONTMATTER_RE = re.compile(r"^---\s*$", re.MULTILINE)


def _parse_frontmatter(md_text: str) -> dict[str, object]:
    import yaml

    matches = list(_FRONTMATTER_RE.finditer(md_text))
    if len(matches) < 2:
        return {}
    start = matches[0].end()
    end = matches[1].start()
    raw = md_text[start:end].strip()
    if not raw:
        return {}
    loaded = yaml.safe_load(raw)
    return loaded if isinstance(loaded, dict) else {}


def _to_decimal(value: object) -> Decimal | None:
    if value is None:
        return None
    s = str(value).strip().replace(",", "")
    if not s:
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    parser.add_argument("--glob", default="**/*.md")
    args = parser.parse_args()

    root = Path(args.path)
    files = [p for p in sorted(root.glob(args.glob)) if p.is_file()]

    required = ["发票种类", "发票号码", "开票日期", "购买方名称", "销售方名称", "金额", "税额", "价税合计"]
    results: list[CheckResult] = []
    for f in files:
        text = f.read_text(encoding="utf-8", errors="replace")
        fm = _parse_frontmatter(text)
        missing = sum(1 for k in required if not str(fm.get(k) or "").strip())

        amount = _to_decimal(fm.get("金额"))
        tax = _to_decimal(fm.get("税额"))
        total = _to_decimal(fm.get("价税合计"))
        mismatch = False
        if amount is not None and tax is not None and total is not None:
            mismatch = (amount + tax) != total

        ok = (missing == 0) and (not mismatch)
        results.append(CheckResult(file=f.name, ok=ok, missing_fields=missing, total_mismatch=mismatch))

    ok_count = sum(1 for r in results if r.ok)
    print(f"checked={len(results)} ok={ok_count} fail={len(results)-ok_count}")
    for r in results:
        if r.ok:
            continue
        flags = []
        if r.missing_fields:
            flags.append(f"missing={r.missing_fields}")
        if r.total_mismatch:
            flags.append("total_mismatch")
        print(f"{r.file}\t" + ",".join(flags))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

