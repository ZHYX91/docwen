"""Built-in Gongwen display values preserve authored YAML."""

from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime

import pytest

from docwen_plugin_markdown.gongwen_fields import prepare_gongwen_fields

pytestmark = pytest.mark.contract


def test_field_projection_keeps_source_data_separate_from_word_layout() -> None:
    source = {
        "附件说明": ["1. 材料清单", "（二）办理依据"],
        "抄送机关": ["省政府办公厅", "市政府办公室"],
        "附注": "（此件公开发布）",
        "印发日期": date(2026, 6, 28),
        "成文日期": "2026/6/27",
        "普通列表": ["甲", ["乙", None]],
    }
    original = deepcopy(source)

    values = prepare_gongwen_fields(source)

    assert source == original
    assert values["附件说明"] == ("材料清单", "办理依据")
    assert values["抄送机关"] == "省政府办公厅，市政府办公室"
    assert values["附注"] == "此件公开发布"
    assert values["印发日期"] == "2026年6月28日印发"
    assert values["成文日期"] == "2026年6月27日"
    assert values["普通列表"] == original["普通列表"]
    assert prepare_gongwen_fields(values) == values


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        (None, ()),
        (["", None, {"nested": "  "}], ()),
        ("附件：材料", ("材料",)),
        (["１． 材料", "(二) 依据", "③ 名单"], ("材料", "依据", "名单")),
        ("１． Ａ２０２６材料", ("Ａ２０２６材料",)),
        ([0, False], ("0", "False")),
        ("2026年度材料", ("2026年度材料",)),
    ],
)
def test_attachment_names_never_contain_layout_padding(source, expected) -> None:
    assert prepare_gongwen_fields({"附件说明": source})["附件说明"] == expected


@pytest.mark.parametrize(
    "source",
    [
        date(2026, 6, 3),
        datetime(2026, 6, 3, 12, 30),
        "2026-06-03",
        "2026/6/3",
        "2026.6.3",
        "2026年6月3日",
        "2026年6月3号",
    ],
)
def test_supported_date_values_have_one_output_format(source) -> None:
    result = prepare_gongwen_fields({"成文日期": source, "印发日期": source})
    assert result == {"成文日期": "2026年6月3日", "印发日期": "2026年6月3日印发"}


@pytest.mark.parametrize("source", ["  待定日期  ", "2026-02-30", "2026年6月3日印发", 0, False])
def test_unparseable_dates_remain_authored_text(source) -> None:
    result = prepare_gongwen_fields({"成文日期": source, "印发日期": source})
    assert result == {"成文日期": str(source), "印发日期": str(source)}


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("（公开）", "公开"),
        ("(公开)", "公开"),
        ("（说明（附页））", "说明（附页）"),
        ("（甲）（乙）", "（甲）（乙）"),
        ("（甲)", "（甲)"),
        (None, ""),
    ],
)
def test_notes_remove_only_one_enclosing_parenthesis_pair(source, expected) -> None:
    assert prepare_gongwen_fields({"附注": source})["附注"] == expected
