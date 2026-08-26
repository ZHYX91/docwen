"""
docx_spell 批注锚点报告工具（命令行）

将带批注的 DOCX（word/document.xml + word/comments.xml）转换为 Markdown 报告，便于人工检查
批注锚点范围与覆盖文本。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docwen_plugin_proofread.anchor_report import build_anchor_report_markdown


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", help="带批注的 DOCX 路径")
    parser.add_argument("--out", help="输出 Markdown 路径；不传则输出到 stdout")
    parser.add_argument("--context-chars", type=int, default=20)
    parser.add_argument("--redact", action="store_true", help="将报告内容做脱敏替换")
    args = parser.parse_args(argv)

    md = build_anchor_report_markdown(args.docx, context_chars=args.context_chars, redact=args.redact)
    if args.out:
        Path(args.out).write_text(md, encoding="utf-8")
        return 0
    print(md, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
