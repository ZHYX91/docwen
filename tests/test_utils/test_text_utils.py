"""utils 单元测试。"""

from __future__ import annotations

import pytest

from docwen.utils.text_utils import clean_text, clean_text_in_data


pytestmark = pytest.mark.unit


@pytest.mark.unit
@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("", ""),
        ("纯文本", "纯文本"),
        ("文本<!-- 这是注释 -->还有内容", "文本还有内容"),
        ("<div>这是<span>嵌套的</span>HTML标签</div>", "这是嵌套的HTML标签"),
        ("空格&nbsp;小于&lt;大于&gt;和号&amp;引号&quot;", '空格 小于<大于>和号&引号"'),
        ("十进制&#20013;&#25991;和十六进制&#x4E2D;&#x6587;", "十进制中文和十六进制中文"),
        ("文本\u200b零\u200c宽\u200d字\ufeff符\u2060测试", "文本零宽字符测试"),
        ("A<br/>B<br />C<br>D", "A\nB\nC\nD"),
    ],
)
def test_clean_text_basic(raw: str, expected: str) -> None:
    assert clean_text(raw) == expected


@pytest.mark.unit
def test_clean_text_in_data_recursive() -> None:
    data = {
        "title": "<b>粗体</b>&nbsp;标题",
        "items": ["<p>项目1</p>", "普通&nbsp;文本"],
        "nested": {"key": "嵌套&lt;内容&gt;"},
    }

    result = clean_text_in_data(data)

    assert result["title"] == "粗体 标题"
    assert result["items"][0] == "项目1"
    assert result["items"][1] == "普通 文本"
    assert result["nested"]["key"] == "嵌套<内容>"
