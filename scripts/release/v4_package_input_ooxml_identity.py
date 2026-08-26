from __future__ import annotations

import hashlib
from collections.abc import Mapping


class V4OoxmlIdentityError(ValueError):
    """A resolved target/reference/citation identity is malformed."""


def _sha(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def target_identity(kind: str, target_id: str) -> tuple[str, str]:
    digest = _sha(f"docwen-target-map-v1\0{kind}\0{target_id}")
    return f"docwen-target-v1:{digest[:32]}", f"DW_T_{digest[:35]}"


def reference_tag(reference: Mapping[str, object], bookmark: str, source_sha256: str) -> str:
    digest = _sha(
        "\0".join(
            (
                "docwen-ref-occurrence-map-v1",
                source_sha256,
                str(reference["source_start"]),
                str(reference["source_end"]),
                str(reference["authored_token"]),
                bookmark,
                str(reference["cached_number"]),
            )
        )
    )
    return f"docwen-ref-occurrence-v1:{digest[:32]}"


def citation_identity(citation: Mapping[str, object], source_sha256: str) -> tuple[str, str, str]:
    raw_items = citation.get("items")
    if not isinstance(raw_items, list) or not raw_items:
        raise V4OoxmlIdentityError("citation_items_invalid")
    refs: list[tuple[str, str, str]] = []
    for raw_item in raw_items:
        if not isinstance(raw_item, dict):
            raise V4OoxmlIdentityError("citation_item_invalid")
        presentation = str(raw_item.get("presentation", ""))
        item_digest = _sha(
            "\0".join(
                (
                    "docwen-citation-item-map-v1",
                    source_sha256,
                    str(raw_item.get("record_id", "")),
                    str(raw_item.get("record_sha256", "")),
                    _sha(presentation),
                )
            )
        )
        word_tag = f"DWCIT_{item_digest[:32]}"
        ref_digest = _sha(
            "\0".join(
                (
                    "docwen-citation-item-ref-v1",
                    str(raw_item.get("citation_key", "")),
                    word_tag,
                    item_digest,
                )
            )
        )
        refs.append((word_tag, item_digest, ref_digest))
    digest = _sha(
        "\0".join(
            (
                "docwen-citation-occurrence-map-v1",
                source_sha256,
                str(citation.get("source_start", "")),
                str(citation.get("source_end", "")),
                str(citation.get("source_slice_sha256", "")),
                str(citation.get("form", "")),
                str(citation.get("cluster_id", "")),
                _sha(str(citation.get("cached_result", ""))),
                ",".join(item[2] for item in refs),
            )
        )
    )
    tokens = ["CITATION", refs[0][0]]
    for word_tag, _item_digest, _ref_digest in refs[1:]:
        tokens.extend((r"\m", word_tag))
    return (
        f"docwen-citation-occurrence-v1:{digest[:32]}",
        f"_DWC_{digest[:35]}",
        f" {' '.join(tokens)} ",
    )


__all__ = ["V4OoxmlIdentityError", "citation_identity", "reference_tag", "target_identity"]
