"""Build YAML metadata from recognition results.

Maps classified paragraphs to the 18 gongwen YAML fields.
Structural element paragraphs are excluded from body rendering (skip_indices).
Field-specific cleaning is applied before writing to yaml_info.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from docwen_plugin_optimizer_gongwen.utils import (
    convert_date_format,
    extract_after_colon,
    extract_combined_id,
    extract_doc_number_and_name,
    extract_doc_number_and_signers,
    extract_printing_line,
    extract_signers_from_text,
    process_attachment_item,
    process_copy_to,
    remove_brackets,
    remove_colon,
)

if TYPE_CHECKING:
    from docwen_plugin_optimizer_gongwen.models import (
        ParagraphFeature,
        RecognitionResult,
    )


# ── Element type → YAML field mapping ─────────────────────────────────
# Maps recognition element types to (yaml_key, is_scalar) pairs.
# is_scalar=True: single string value (don't overwrite).
# is_scalar=False: list value (append).
_ELEMENT_TO_YAML: dict[str, tuple[str, bool]] = {
    "title": ("标题", True),
    "subtitle": ("副标题", True),
    "copy_id": ("份号", True),
    "security": ("密级和保密期限", True),
    "urgency": ("紧急程度", True),
    "doc_number": ("发文字号", True),
    "issuing_authority_mark": ("发文机关标志", True),
    "signer": ("签发人", False),
    "issuing_authority_signature": ("发文机关署名", True),
    "issue_date": ("成文日期", True),
    "printing_date": ("印发日期", True),
    "recipient": ("主送机关", True),
    "notes": ("附注", True),
    "printing_authority": ("印发机关", True),
    "copy_to": ("抄送机关", False),
    "attachment_header": ("附件说明", False),
    "disclosure": ("公开方式", True),
}


def _looks_like_attachment_list_item(text: str) -> bool:
    """Return True for attachment-list continuations misclassified as body content."""
    return bool(re.match(r"^\s*(?:\d+[.．]|[（(]?[一二三四五六七八九十]+[）)、.．、])\s*", text))


def build_yaml(
    scorer,
    features: list[ParagraphFeature],
    result: RecognitionResult,
    *,
    cleanup_rules=(),
) -> RecognitionResult:
    """Convert recognition candidates to YAML metadata.

    Maps classified paragraphs to the 18 gongwen YAML fields.
    Structural element paragraphs are added to skip_indices.
    """
    yaml_info: dict[str, str | list[str]] = {
        "aliases": [],
        "标题": "",
        "副标题": "",
        "份号": "",
        "密级和保密期限": "",
        "紧急程度": "",
        "发文字号": "",
        "发文机关标志": "",
        "签发人": [],
        "发文机关署名": "",
        "成文日期": "",
        "印发日期": "",
        "主送机关": "",
        "附注": "",
        "印发机关": "",
        "抄送机关": [],
        "附件说明": [],
        "公开方式": "",
    }

    skip_indices: list[int] = []
    attachment_list_indices: list[int] = []

    for para_idx, candidate in result.candidates.items():
        text = features[para_idx].text.strip() if para_idx < len(features) else ""
        element_type = candidate.element_type
        yaml_key: str | None = None
        is_scalar: bool = True

        # ── Composite fields ──────────────────────────────────────
        if element_type == "combined_id":
            copy_id, doc_number = extract_combined_id(text)
            if copy_id and not yaml_info["份号"]:
                yaml_info["份号"] = copy_id
            if doc_number and not yaml_info["发文字号"]:
                yaml_info["发文字号"] = doc_number
            skip_indices.append(para_idx)
            continue

        if element_type == "combined_doc_number_signer":
            doc_number, signers = extract_doc_number_and_signers(text)
            if doc_number and not yaml_info["发文字号"]:
                yaml_info["发文字号"] = doc_number
            signer_values = yaml_info["签发人"]
            if isinstance(signer_values, list):
                for signer in signers:
                    if signer not in signer_values:
                        signer_values.append(signer)
            skip_indices.append(para_idx)
            continue

        if element_type == "combined_doc_number_signer_following":
            doc_number, signer = extract_doc_number_and_name(text)
            if doc_number and not yaml_info["发文字号"]:
                yaml_info["发文字号"] = doc_number
            signer_values = yaml_info["签发人"]
            if signer and isinstance(signer_values, list) and signer not in signer_values:
                signer_values.append(signer)
            skip_indices.append(para_idx)
            continue

        if element_type == "signer":
            signer_values = yaml_info["签发人"]
            if isinstance(signer_values, list):
                for signer in extract_signers_from_text(extract_after_colon(text)):
                    if signer not in signer_values:
                        signer_values.append(signer)
            skip_indices.append(para_idx)
            continue

        # ── Direct mapping ──
        if element_type in _ELEMENT_TO_YAML:
            yaml_key, is_scalar = _ELEMENT_TO_YAML[element_type]

        # ── Title following → contribute to title ──
        elif element_type == "title_following":
            if text:
                cleaned = text.replace("\n", "").replace("\r", "")
                if cleaned:
                    assert isinstance(yaml_info["标题"], str)
                    yaml_info["标题"] += cleaned
            skip_indices.append(para_idx)
            continue

        # ── Attachment list continuation → 附件说明 ──
        elif element_type == "attachment_following":
            yaml_key, is_scalar = "附件说明", False
            attachment_list_indices.append(para_idx)

        # ── Attachment body content stays out of YAML metadata ────────
        elif element_type == "attachment_content":
            if _looks_like_attachment_list_item(text):
                yaml_key, is_scalar = "附件说明", False
                attachment_list_indices.append(para_idx)
            else:
                skip_indices.append(para_idx)
                continue
        # ── Subtitle following → contribute to subtitle ──
        elif element_type == "subtitle_following":
            if text:
                cleaned = text.replace("\n", "").replace("\r", "")
                if cleaned:
                    assert isinstance(yaml_info["副标题"], str)
                    yaml_info["副标题"] += cleaned
            skip_indices.append(para_idx)
            continue

        # ── Signer following → contribute to signer list ──
        elif element_type == "signer_following":
            if text:
                extracted = extract_signers_from_text(text)
                if extracted:
                    assert isinstance(yaml_info["签发人"], list)
                    for s in extracted:
                        if s not in yaml_info["签发人"]:
                            yaml_info["签发人"].append(s)
            skip_indices.append(para_idx)
            continue

        # ── Non-structural content → no YAML mapping ──
        elif element_type in ("body",):
            # These don't produce top-level YAML entries
            continue

        else:
            # Unknown type, skip
            continue

        if yaml_key is None or not text:
            continue

        # ── Apply field-specific cleaning (Task 4) ──
        cleaned_text = text
        if element_type == "issue_date":
            cleaned_text = convert_date_format(text)
        elif element_type == "printing_date":
            printing_authority, printing_date = extract_printing_line(text)
            if printing_authority and not yaml_info["印发机关"]:
                yaml_info["印发机关"] = printing_authority
            cleaned_text = printing_date or convert_date_format(re.sub(r"\s*印发\s*$", "", text))
        elif element_type == "recipient":
            cleaned_text = remove_colon(text)
        elif element_type == "notes":
            cleaned_text = remove_brackets(text)
        elif element_type == "disclosure":
            cleaned_text = extract_after_colon(text)
        elif element_type == "copy_to":
            yaml_info["抄送机关"] = process_copy_to(text)
            # Skip scalar/list assignment below — we directly assign a list
            skip_indices.append(para_idx)
            continue
        elif element_type in ("attachment_header", "attachment_following", "attachment_content"):
            cleaned_text = process_attachment_item(
                text,
                cleanup_rules=cleanup_rules,
            )

        if is_scalar:
            # Don't overwrite an already-assigned scalar field
            if not yaml_info[yaml_key]:
                yaml_info[yaml_key] = cleaned_text
        else:
            current = yaml_info[yaml_key]
            if isinstance(current, list) and cleaned_text not in current:
                current.append(cleaned_text)

        # Structural paragraphs are excluded from body rendering
        if element_type not in ("body",):
            skip_indices.append(para_idx)

    # ── Extract clean document number from combined fields ──
    if yaml_info["发文字号"]:
        doc_num = yaml_info["发文字号"]
        m = re.search(
            r"([A-Za-z一-鿿]+[〔\(\[【]\d{4}[〕\)\]】]\d+号?)",
            str(doc_num),
        )
        if m:
            yaml_info["发文字号"] = m.group(1)

    title = str(yaml_info["标题"]).strip()
    aliases = yaml_info["aliases"]
    if title and isinstance(aliases, list) and title not in aliases:
        aliases.append(title)

    result.yaml_info = yaml_info
    result.skip_indices = list(set(skip_indices))

    return result
