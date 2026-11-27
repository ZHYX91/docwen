"""Imported organizations retain punctuation; signers are never silently discarded."""

from __future__ import annotations

import pytest

from docwen_plugin_optimizer_gongwen.models import ParagraphFeature, RecognitionCandidate, RecognitionResult
from docwen_plugin_optimizer_gongwen.recognition.signals import collect_structured_signals
from docwen_plugin_optimizer_gongwen.recognition.yaml_builder import build_yaml
from docwen_plugin_optimizer_gongwen.utils import extract_signers_from_text, process_copy_to

pytestmark = pytest.mark.contract


def _build(texts: list[str], kinds: list[str]):
    features = [ParagraphFeature(index=index, text=text, source_index=index) for index, text in enumerate(texts)]
    result = RecognitionResult(
        candidates={
            index: RecognitionCandidate(element_type=kind, score=100, para_index=index)
            for index, kind in enumerate(kinds)
        },
        yaml_info={},
        skip_indices=[],
        review_signals=[],
        missing_required=[],
        validation_finding_count=0,
    )
    return features, build_yaml(None, features, result)


@pytest.mark.parametrize(
    "text,expected",
    [
        ("抄送：各县（市、区）人民政府，省直各单位。", "各县（市、区）人民政府，省直各单位"),
        ("报送：甲局、乙局；丙单位。", "甲局、乙局；丙单位"),
        ("分送: 甲（乙,丙）单位，丁单位.", "甲（乙,丙）单位，丁单位"),
        ("甲局、乙局，丙单位", "甲局、乙局，丙单位"),
    ],
)
def test_copy_recipients_keep_the_whole_authored_expression(text, expected) -> None:
    assert process_copy_to(text) == expected
    _features, result = _build([text], ["copy_to"])
    assert result.yaml_info["抄送机关"] == expected


@pytest.mark.parametrize(
    "text,expected",
    [
        ("张三、李四", ["张三", "李四"]),
        ("张三，李四;王五", ["张三", "李四", "王五"]),
        ("张三 李四", ["张三", "李四"]),
        ("张　三", ["张　三"]),
        ("张 三　李四", ["张 三　李四"]),
        ("𠮷田、阿里·艾力", ["𠮷田", "阿里·艾力"]),
        ("John Smith\t王五", ["John Smith", "王五"]),
        ("签发人：张三、张三", ["张三", "张三"]),
        ("单字名：甲", ["单字名：甲"]),
        ("签发人：", []),
    ],
)
def test_signers_preserve_unusual_names_ambiguous_spacing_and_duplicates(text, expected) -> None:
    assert extract_signers_from_text(text) == expected


def test_signer_paragraphs_are_processed_in_source_order_without_deduplication() -> None:
    features, result = _build(["签发人：张三", "李四、张三"], ["signer", "signer_following"])
    result.candidates = dict(reversed(list(result.candidates.items())))
    result = build_yaml(None, features, result)
    assert result.yaml_info["签发人"] == ["张三", "李四", "张三"]
    assert set(result.skip_indices) == {0, 1}


def test_empty_signer_extraction_keeps_the_source_paragraph_and_requires_review() -> None:
    features, result = _build(["签发人："], ["signer"])
    assert result.yaml_info["签发人"] == []
    assert 0 not in result.skip_indices
    assert "signer_text_preserved" in result.candidates[0].trace

    signals = collect_structured_signals(result, features)
    assert signals["recognition_summary"]["needs_review"]
    assert any(warning["code"] == "GW003" for warning in signals["gongwen_warnings"])


def test_ambiguous_spaced_name_is_retained_in_yaml_and_marked_for_review() -> None:
    _features, result = _build(["签发人：张　三"], ["signer"])
    assert result.yaml_info["签发人"] == ["张　三"]
    assert "signer_text_preserved" in result.candidates[0].trace


def test_recipient_continuation_keeps_mixed_punctuation_and_consumes_both_paragraphs() -> None:
    _features, result = _build(
        ["各县（市、区）人民政府，", "省直各单位："],
        ["body", "recipient"],
    )
    assert result.yaml_info["主送机关"] == "各县（市、区）人民政府，省直各单位"
    assert set(result.skip_indices) == {0, 1}


@pytest.mark.parametrize("earlier", ["请通知各单位，", "会议要求各单位，", "普通正文。", "未明确的前文，"])
def test_recipient_does_not_absorb_ordinary_or_uncertain_previous_text(earlier) -> None:
    _features, result = _build([earlier, "省直各单位："], ["body", "recipient"])
    assert result.yaml_info["主送机关"] == "省直各单位"
    assert 0 not in result.skip_indices


def test_recipient_does_not_cross_a_blank_source_gap_or_table_boundary() -> None:
    features, result = _build(["各市人民政府，", "省直各单位："], ["body", "recipient"])
    features[1].source_index = 3
    result = build_yaml(None, features, result)
    assert result.yaml_info["主送机关"] == "省直各单位"
    assert 0 not in result.skip_indices
    features[1].source_index = 1
    features[0].source = "table"
    result = build_yaml(None, features, result)
    assert result.yaml_info["主送机关"] == "省直各单位"
    assert 0 not in result.skip_indices


@pytest.mark.parametrize(
    "attribute", ["has_image", "has_formula", "has_page_break", "has_section_break", "heading_level", "list_marker"]
)
def test_recipient_does_not_consume_objects_breaks_headings_or_lists(attribute) -> None:
    features, result = _build(["各市人民政府，", "省直各单位："], ["body", "recipient"])
    setattr(features[0], attribute, True)
    result = build_yaml(None, features, result)
    assert result.yaml_info["主送机关"] == "省直各单位"
    assert 0 not in result.skip_indices


@pytest.mark.parametrize(
    "earlier",
    ["承办单位：省财政厅，", "办理联系人为省财政厅，", "工作由省财政厅，", "省财政厅，", "各单位应联系省财政厅，"],
)
def test_recipient_keeps_uncertain_agency_suffix_prose_and_requests_review(earlier: str) -> None:
    from docwen_plugin_optimizer_gongwen.recognition.rounds import run_rounds
    from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

    texts = ["关于测试的通知", earlier, "各市人民政府：", "正文内容。"]
    features = [
        ParagraphFeature(
            index=i,
            text=text,
            source_index=i,
            font_name="小标宋" if i == 0 else "仿宋",
            font_size_pt=22 if i == 0 else 16,
        )
        for i, text in enumerate(texts)
    ]
    scorer = ElementScorer()
    result = build_yaml(scorer, features, run_rounds(scorer, features))

    assert result.yaml_info["主送机关"] == "各市人民政府"
    assert 1 not in result.skip_indices
    signals = collect_structured_signals(result, features)
    assert any(warning["code"] == "GW004" for warning in signals["gongwen_warnings"])


@pytest.mark.parametrize("names", ["𠮷田", "John Smith", "𠮷田、John Smith、张三、张三"])
def test_labelled_composite_signers_survive_the_full_docx_pipeline(names: str) -> None:
    from docx import Document
    from docx.shared import Pt

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    first = document.add_paragraph(f"某发〔2026〕1号\t签发人：{names}")
    first.runs[0].font.name = "仿宋"
    first.runs[0].font.size = Pt(16)
    title = document.add_paragraph("关于测试的通知")
    title.runs[0].font.name = "小标宋"
    title.runs[0].font.size = Pt(22)
    document.add_paragraph("各市人民政府：")
    document.add_paragraph("正文内容。")

    result = convert_docx_to_md_gongwen(document, "signers.docx", options={})

    assert result["success"]
    assert result["yaml_info"]["发文字号"] == "某发〔2026〕1号"
    assert result["yaml_info"]["签发人"] == names.split("、")
    assert "正文内容。" in result["markdown"]


@pytest.mark.parametrize("text", ["某发〔2026〕1号 正文内容", "正文内容 签发人：John Smith"])
def test_unlabelled_or_invalid_document_number_prose_is_not_a_composite_signer(text: str) -> None:
    from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

    assert not ElementScorer().has_doc_number_and_signer(ParagraphFeature(index=0, text=text))


@pytest.mark.parametrize(
    "earlier",
    ["各县（市、区）人民政府，", "省直各单位，", "各应急管理局，", "各自由贸易试验区办公室，", "各教育工作委员会，"],
)
def test_clear_collective_address_continuations_keep_legitimate_agency_names(earlier: str) -> None:
    _features, result = _build([earlier, "各市人民政府："], ["body", "recipient"])
    assert result.yaml_info["主送机关"] == earlier + "各市人民政府"
    assert 0 in result.skip_indices


def test_uncertain_recipient_prose_remains_in_the_full_docx_markdown_output() -> None:
    from docx import Document
    from docx.shared import Pt

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    title = document.add_paragraph("关于测试的通知")
    title.runs[0].font.name = "小标宋"
    title.runs[0].font.size = Pt(22)
    document.add_paragraph("承办单位：省财政厅，")
    document.add_paragraph("各市人民政府：")
    document.add_paragraph("正文内容。")

    result = convert_docx_to_md_gongwen(document, "recipient.docx", options={})

    assert result["yaml_info"]["主送机关"] == "各市人民政府"
    assert "承办单位：省财政厅，" in result["markdown"]
    assert "正文内容。" in result["markdown"]
    signals = result["metadata"]["recognition_review_signals"]
    assert any(warning["code"] == "GW004" for warning in signals["gongwen_warnings"])


@pytest.mark.parametrize(
    "line,expected",
    [
        ("各有关单位：", "各有关单位"),
        ("示例研发中心：", "示例研发中心"),
        ("各区（A区、B区）：", "各区（A区、B区）"),
        ("示例工信委：", "示例工信委"),
        ("各县（市、区）人民政府，示例研发中心:", "各县（市、区）人民政府，示例研发中心"),
        ("各位同仁：", "各位同仁"),
        ("罗代表：", "罗代表"),
        ("主送机关：示例研发中心", ""),
        ("各有关部门参加了讨论，记录如下：", ""),
        ("各部门应将书面材料提交示例研发中心：", ""),
        ("讨论已结束。经各单位确认，现将决定列举如下：", ""),
    ],
)
def test_recipient_recognizes_an_opening_address(line: str, expected: str) -> None:
    from docx import Document

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    document.add_paragraph("示例事项通知")
    document.add_paragraph(line)
    document.add_paragraph("正文开始。")

    result = convert_docx_to_md_gongwen(document, "synthetic.docx", options={})

    assert result["yaml_info"]["主送机关"] == expected
    body = result["markdown"].split("---", 2)[2]
    assert (line in body) == (not expected)


@pytest.mark.parametrize("body_before_address", [False, True])
def test_recipient_stops_at_the_body_boundary(body_before_address: bool) -> None:
    from docx import Document
    from docx.shared import Pt

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    document.add_paragraph("示例事项通知")
    if body_before_address:
        document.add_paragraph("以下列出本次工作的详细情况。")
    address = document.add_paragraph("示例研发中心：")
    if not body_before_address:
        address.paragraph_format.first_line_indent = Pt(28)
    document.add_paragraph("正文开始。")

    result = convert_docx_to_md_gongwen(document, "synthetic.docx")

    assert result["yaml_info"]["主送机关"] == ""
    assert address.text in result["markdown"].split("---", 2)[2]


@pytest.mark.parametrize("opening", ["title_continuation", "centered_note", "indented_list"])
def test_recipient_uses_the_full_opening_layout(opening: str) -> None:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Pt

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    title = document.add_paragraph("关于开展试点的通知")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if opening == "title_continuation":
        lines = [title, document.add_paragraph("第一批示范事项"), document.add_paragraph("联动实施安排")]
        for line in lines:
            run = line.runs[0]
            run.font.name = "方正小标宋"
            run.font.size = Pt(22)
            run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "方正小标宋")
    elif opening == "centered_note":
        note = document.add_paragraph("（2026年1月19日）")
        note.alignment = WD_ALIGN_PARAGRAPH.CENTER

    addressees = "示例协作平台、各位观察员" if opening == "indented_list" else "各位观察员"
    address = document.add_paragraph(f"{addressees}：")
    body = document.add_paragraph("为做好试点工作，现将有关事项说明如下：")
    body.paragraph_format.first_line_indent = Pt(28)
    if opening == "indented_list":
        address.paragraph_format.first_line_indent = Pt(28)
    document.add_paragraph("试点各方应建立工作台账。")

    result = convert_docx_to_md_gongwen(document, "synthetic.docx")

    assert result["yaml_info"]["主送机关"] == addressees
    markdown_body = result["markdown"].split("---", 2)[2]
    assert address.text not in markdown_body
    assert body.text in markdown_body


@pytest.mark.parametrize("number_before_title", [False, True])
def test_meeting_number_and_typed_decisions_keep_their_source_content(number_before_title: bool) -> None:
    from docx import Document

    from docwen_core.text.heading_numbering import compile_clean_rules_from_data
    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    title = "示例演练会议纪要"
    number = "（2038）7号"
    narrative = "示例小组查看了各部门的演练记录。有关决定如下："
    unclear = "示例演练联络组　2038年3月12日"
    body_reference = "(2037) 6号"
    for line in (number, title) if number_before_title else (title, number):
        document.add_paragraph(line)
    document.add_paragraph(narrative)
    document.add_paragraph("1.继续检查设备。记录完成时间。")
    document.add_paragraph("2.复核演练程序。保留变更记录。")
    document.add_paragraph("3.复核周期。", style="Heading 3")
    document.add_paragraph("供复核的另一份材料：")
    document.add_paragraph(body_reference)
    document.add_paragraph(unclear)

    result = convert_docx_to_md_gongwen(
        document,
        "synthetic.docx",
        options={},
        cleanup_rules=compile_clean_rules_from_data(
            [
                {"id": "arabic_separator", "level": 3, "pattern": r"^\d+\."},
                {"id": "arabic_bracket", "level": 4, "pattern": r"^[（(]\d+[）)]"},
            ]
        ),
    )

    assert result["yaml_info"]["发文字号"] == number
    assert result["yaml_info"]["主送机关"] == ""
    assert result["yaml_info"]["发文机关署名"] == ""
    assert result["yaml_info"]["成文日期"] == ""
    body = result["markdown"].split("---", 2)[2]
    assert number not in body
    assert narrative in body
    assert "1. 继续检查设备。记录完成时间。" in body
    assert "2. 复核演练程序。保留变更记录。" in body
    assert "### 复核周期。" in body
    assert body_reference in body
    assert unclear in body


def test_gongwen_body_preserves_accepted_inline_revisions() -> None:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    document.add_paragraph("关于办理事项的通知")
    document.add_paragraph("各学院：")
    body = document.add_paragraph("请于周五报送")
    inserted = OxmlElement("w:ins")
    inserted.set(qn("w:id"), "1")
    inserted.append(body.add_run("。")._r)
    body._p.append(inserted)
    deleted = OxmlElement("w:del")
    deleted.set(qn("w:id"), "2")
    removed_run = OxmlElement("w:r")
    removed_text = OxmlElement("w:delText")
    removed_text.text = "旧时间"
    removed_run.append(removed_text)
    deleted.append(removed_run)
    body._p.append(deleted)
    body.add_run("谢谢配合。")

    result = convert_docx_to_md_gongwen(document, "synthetic.docx")
    assert result["yaml_info"]["主送机关"] == "各学院"
    assert "请于周五报送。谢谢配合。" in result["markdown"]
    assert "旧时间" not in result["markdown"]


def test_gongwen_attachment_renders_each_table_cell_once() -> None:
    from docx import Document

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    document.add_paragraph("关于材料归档的通知")
    document.add_paragraph("各学院：")
    document.add_paragraph("请按随文附表归档材料。")
    document.add_paragraph("附件：归档材料表")
    document.add_paragraph("示例地区人民政府")
    document.add_paragraph("2026年10月20日")
    document.add_paragraph("归档材料表")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "材料类型"
    table.cell(0, 1).text = "办理要求"
    table.cell(1, 0).text = "书面计划"
    table.cell(1, 1).text = "说明资料提交方式"

    result = convert_docx_to_md_gongwen(document, "synthetic.docx")
    (attachment,) = result["attachment_documents"]
    assert "| 书面计划 | 说明资料提交方式 |" in attachment.markdown
    assert attachment.markdown.count("说明资料提交方式") == 1


@pytest.mark.parametrize("number_before_title", [False, True])
def test_meeting_header_signature_and_date_are_projected_together(number_before_title: bool) -> None:
    from docx import Document

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    title = "示例工作会议纪要"
    number = "（2036）4号"
    signature = "演练管理局"
    issue_date = "2036年2月29日"
    signature_line = f"{signature}          {issue_date}"
    lead = "当天召集各部门会商。有关决定如下："
    body_pair = "示例联络组　2036年3月1日"
    body_number = "(2034) 8号"
    decision = "按计划核对设备。"

    document = Document()
    for line in (number, title) if number_before_title else (title, number):
        document.add_paragraph(line)
    document.add_paragraph(signature_line)
    document.add_paragraph(lead)
    document.add_paragraph(decision)
    document.add_paragraph(body_number)
    document.add_paragraph(body_pair)
    document.add_paragraph("保留上述材料供复查。")

    result = convert_docx_to_md_gongwen(document, "synthetic.docx", options={})

    assert result["yaml_info"]["发文字号"] == number
    assert result["yaml_info"]["发文机关署名"] == signature
    assert result["yaml_info"]["成文日期"] == issue_date
    assert result["yaml_info"]["主送机关"] == ""
    assert result["metadata"]["missing_required"] == []
    confidence = result["metadata"]["confidence"]
    assert confidence["finding_count"] == 0
    assert {"issuing_authority_signature", "issue_date", "body"} <= confidence["fields"].keys()
    signals = result["metadata"]["recognition_review_signals"]
    assert signals["used_fallback"] == []
    body = result["markdown"].split("---", 2)[2]
    assert signature_line not in body
    assert lead in body and decision in body
    assert body_number in body and body_pair in body


@pytest.mark.parametrize(
    "signature_line, lead",
    [
        ("示例管理局　2035年2月29日", "会商结束。有关事项如下："),
        ("示例管理局\n2035年2月28日", "会商结束。有关事项如下："),
        ("示例管理局 2035年2月28日", "一、参会名单"),
        ("负责办理的业务 2035年2月28日", "会商结束。有关事项如下："),
    ],
)
def test_unproven_meeting_header_pair_remains_visible(signature_line: str, lead: str) -> None:
    from docx import Document

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    for line in ("示例会商纪要", "（2035）9号", signature_line, lead, "其余材料继续留存。"):
        document.add_paragraph(line)
    result = convert_docx_to_md_gongwen(document, "synthetic.docx", options={})

    assert result["yaml_info"]["发文机关署名"] == ""
    assert result["yaml_info"]["成文日期"] == ""
    assert set(result["metadata"]["missing_required"]) == {"issuing_authority_signature", "issue_date"}
    assert signature_line in result["markdown"].split("---", 2)[2]


def test_meeting_signature_and_date_with_page_break_remain_in_body() -> None:
    from docx import Document
    from docx.enum.text import WD_BREAK

    from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

    document = Document()
    document.add_paragraph("示例工作会议纪要")
    document.add_paragraph("（2036）4号")
    pair = "演练管理局　2036年2月29日"
    document.add_paragraph(pair).add_run().add_break(WD_BREAK.PAGE)
    document.add_paragraph("当天召集各部门会商。有关决定如下：")
    result = convert_docx_to_md_gongwen(document, "synthetic.docx")

    assert result["yaml_info"]["发文机关署名"] == ""
    assert result["yaml_info"]["成文日期"] == ""
    assert pair in result["markdown"].split("---", 2)[2]
