"""Generate a deterministic Explorer drag/drop smoke fixture."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("tmp") / "explorer-dragdrop-smoke-fixture",
        help="Directory where fixture files and manifest.json are written.",
    )
    parser.add_argument(
        "--supported-count",
        type=int,
        default=245,
        help="Number of supported Markdown files to generate.",
    )
    parser.add_argument(
        "--unsupported-count",
        type=int,
        default=7,
        help="Number of unsupported sidecar files to generate.",
    )
    args = parser.parse_args(argv)

    if args.supported_count < 0 or args.unsupported_count < 0:
        parser.error("counts must be non-negative")

    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    supported_files: list[str] = []
    for index in range(1, args.supported_count + 1):
        name = f"docwen-dragdrop-{index:03d}.md"
        (output_dir / name).write_text(
            f"# Dragdrop {index}\n\nGenerated for DocWen Explorer drag/drop smoke.\n",
            encoding="utf-8",
        )
        supported_files.append(name)

    unsupported_files: list[str] = []
    for index in range(1, args.unsupported_count + 1):
        name = f"skip-{index:02d}.unsupported"
        (output_dir / name).write_text(
            "unsupported fixture file\n",
            encoding="utf-8",
        )
        unsupported_files.append(name)

    manifest = {
        "fixture": "explorer_dragdrop_smoke",
        "output_dir": str(output_dir),
        "supported_count": args.supported_count,
        "unsupported_count": args.unsupported_count,
        "total_count": args.supported_count + args.unsupported_count,
        "supported_extension": ".md",
        "unsupported_extension": ".unsupported",
        "supported_files_preview": supported_files[:5],
        "unsupported_files_preview": unsupported_files[:5],
    }
    (output_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(manifest, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
