from __future__ import annotations

import argparse
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    parser.add_argument("--glob", default="**/*.pdf")
    args = parser.parse_args()

    import fitz

    root = Path(args.path)
    files = sorted(root.glob(args.glob))
    for f in files:
        if not f.is_file():
            continue
        try:
            doc = fitz.open(str(f))
        except Exception:
            print(f"{f.name}\topen_failed")
            continue

        try:
            lens: list[int] = []
            for page in doc:
                t = page.get_text("text") or ""
                lens.append(len(t.strip()))
            print(f"{f.name}\tpages={len(lens)}\ttext_lens={lens}")
        finally:
            doc.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
