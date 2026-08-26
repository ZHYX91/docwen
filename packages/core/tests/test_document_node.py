from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from docwen_core.models import (
    ConversionIdentity,
    DocumentNodePath,
    DocumentNodeValidationError,
    canonical_source_tag,
    sanitize_node_label,
    validate_markdown_node_path,
)

pytestmark = pytest.mark.unit


def test_conversion_identity_uses_one_timezone_aware_timestamp() -> None:
    identity = ConversionIdentity.create(
        task_id="task.1",
        source_stem="公文",
        source_format="docx",
        created_at=datetime(2026, 8, 20, 21, 45, tzinfo=timezone(timedelta(hours=8))),
    )

    assert identity.node_name() == "公文_20260820_214500_fromDocx"
    assert identity.node_name("公文_附件") == "公文_附件_20260820_214500_fromDocx"
    assert identity.created_at_utc == "2026-08-20T13:45:00Z"


def test_document_node_path_enforces_matching_parent_and_markdown_stem() -> None:
    root = DocumentNodePath("Report_20260820_214500_fromDocx")
    child = root.child("Report_附件_20260820_214500_fromDocx")

    assert root.markdown.as_posix().endswith("Report_20260820_214500_fromDocx/Report_20260820_214500_fromDocx.md")
    assert child.markdown.parent.name == child.markdown.stem
    validate_markdown_node_path(child.markdown.as_posix())

    with pytest.raises(DocumentNodeValidationError):
        validate_markdown_node_path("Report/leaf.md")


def test_node_label_is_portable_and_bounded() -> None:
    assert sanitize_node_label(" CON ") == "_CON"
    assert sanitize_node_label("a/b:c") == "a_b_c"
    assert len(sanitize_node_label("很长" * 100)) <= 96
    assert canonical_source_tag(".MHTML") == "Mhtml"


def test_naive_conversion_time_is_rejected() -> None:
    with pytest.raises(DocumentNodeValidationError):
        ConversionIdentity.create(
            task_id="task.1",
            source_stem="report",
            source_format="docx",
            created_at=datetime(2026, 8, 20, 21, 45),
        )
