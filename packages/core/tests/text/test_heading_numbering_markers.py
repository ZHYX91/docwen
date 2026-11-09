"""Number Suite provenance compatibility, independent of cleanup schemes."""

import pytest

from docwen_core.text.heading_numbering import (
    HeadingFormatter,
    compile_clean_rules_from_data,
    detect_heading_prefix,
    strip_heading_prefix,
)

pytestmark = pytest.mark.unit
WJ = "\u2060"


@pytest.mark.parametrize("label", ["1.", "1.2.3", "第一章", "（一）", "①", "IV.", "Appendix A:"])
def test_complete_marker_is_authoritative(label):
    text = f"{WJ}{label}{WJ}\t  Title"
    info = detect_heading_prefix(text, rules=())
    assert info is not None and info.rule_id == "plugin-marker"
    assert strip_heading_prefix(text, rules=()) == (f"{WJ}{label}{WJ}\t  ", "Title")


@pytest.mark.parametrize(
    "text",
    [f"{WJ}1. Title", f"1. Title{WJ}", f"{WJ}{WJ}Title", f"{WJ} \t{WJ}Title", f"{WJ}1.\n2.{WJ}Title", f"{WJ}1.{WJ}\t "],
)
def test_incomplete_or_unsafe_marker_is_preserved(text):
    rules = compile_clean_rules_from_data([{"id": "manual", "pattern": r"^1\. ", "level": 1}])
    assert strip_heading_prefix(text, rules=rules) == ("", text)


@pytest.mark.parametrize("separator", ["", " ", "\t", "\u3000", "\u00a0", "  \t", "\u3000\u00a0 "])
def test_generated_marker_wraps_full_label_not_separator(separator):
    formatter = HeadingFormatter(
        {"level_1": {"format": "第{1.chinese_lower}章" + separator}}, mark_generated_numbers=True
    )
    text = formatter.format_heading("Title", 1)
    assert text == f"{WJ}第一章{WJ}{separator}Title"
    assert strip_heading_prefix(text, rules=()) == (f"{WJ}第一章{WJ}{separator}", "Title")


def test_default_off_and_existing_marker_never_duplicated():
    formatter = HeadingFormatter({"level_1": {"format": "{1.arabic_half}. "}})
    assert formatter.format_heading("Title", 1) == "1. Title"
    for title in [f"{WJ}9.{WJ} Title", f"{WJ}9. Title", f"Title{WJ}"]:
        assert formatter.format_heading(title, 1) == title
    assert formatter.format_heading("Next", 1) == "5. Next"
