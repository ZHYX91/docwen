"""Gongwen field constants, round groups, thresholds, and regex patterns."""

from __future__ import annotations

import re

# ── 18 field names ────────────────────────────────────────────────────
FIELDS = [
    "aliases",
    "标题",
    "副标题",
    "份号",
    "密级和保密期限",
    "紧急程度",
    "发文字号",
    "发文机关标志",
    "签发人",
    "发文机关署名",
    "成文日期",
    "印发日期",
    "主送机关",
    "附注",
    "印发机关",
    "抄送机关",
    "附件说明",
    "公开方式",
]

# Ordered recognition rounds; stable ordering breaks equal-score ties.
ROUND1_ELEMENTS = (
    "combined_id",
    "copy_id",
    "security",
    "urgency",
    "doc_number",
    "combined_doc_number_signer",
    "signer",
    "title",
    "recipient",
    "attachment_header",
    "issuing_authority_signature",
    "issue_date",
    "notes",
    "disclosure",
    "copy_to",
    "printing_date",
)

ROUND2_ELEMENTS = (
    "issuing_authority_mark",
    "printing_authority",
)

ROUND3_ELEMENTS = (
    "title_following",
    "subtitle",
    "subtitle_following",
    "body",
    "attachment_following",
    "signer_following",
    "combined_doc_number_signer_following",
    "attachment_content",
)

# Thresholds
UNIQUE_SCORE_THRESHOLD = 80
NON_UNIQUE_SCORE_THRESHOLD = 60

# Font heuristics
TITLE_FONTS = ["小标宋", "方正小标宋", "华文中宋"]
BODY_FONTS = ["仿宋", "仿宋_GB2312", "华文仿宋"]
HEITI_FONTS = ["黑体", "华文黑体"]

# Regex patterns
RE_DOC_NUMBER = re.compile(
    r"[A-Za-z一-鿿]+[发办函通示件]?"
    r"[〔《【［［（（〔\[]\d{4}[〕》】］］））〕\]]"
    r"\d+[号号＃ ]*"
)
RE_DATE = re.compile(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日")
RE_PRINTING_DATE = re.compile(r"(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)\s*[印印发]")
RE_SECURITY = re.compile(r"(秘密|机密|绝密)")
RE_COPY_ID = re.compile(r"^\s*0*\d+\s*$")
COPY_TO_LABELS = ("抄送", "报送", "分送")
RE_ATTACHMENT_HEADER = re.compile(r"(附件|附件说明|附件：|附件\d+|附件\d+：)")
RE_RECIPIENT = re.compile(r"^[各省市区县自治一-鿿、，；\s]+[：:]")
URGENCY_KEYWORDS = ["特急", "加急", "平急"]
