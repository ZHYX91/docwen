# MD→DOCX Golden Baseline (旧系统单向输出)

**生成日期:** 2026-06-12
**来源:** 旧项目 `S:/OneDrive/Projects/docwen旧/` @ feat-pyside6-gui 同期快照
**用途:** 收尾阶段 §9 第 0 步前置基线，固定旧系统 MD→DOCX 行为作为回归对照
**目录:** `tests/fixtures/golden/md_to_docx_old/`

## 生成方式

```python
# 旧系统 MD→DOCX 单向转换
from docwen.converter.md.to_docx import convert_md_to_docx
convert_md_to_docx(
    md_path="samples/sample.md",
    output_path="sample_golden.docx",
    template_name="简体中文通用模板",
)
```

## 文件说明

| 文件 | 说明 | 对应 Golden ID |
|------|------|---------------|
| `sample_golden.docx` | 旧系统 MD→DOCX 单向输出（输入: `samples/sample.md`） | GOLDEN-001b |
| `sample_roundtrip.md` | 旧系统 MD→DOCX→MD 往返输出（将 sample_golden.docx 作为 DOCX→MD 输入） | GOLDEN-001 old_output_path |

## 比较策略

- **GOLDEN-001b**: 收尾后新系统 MD→DOCX 单向输出应与 `sample_golden.docx` 进行段落文本/标题层级/表格结构/图片数量语义比较
- **GOLDEN-001 old_output**: 用于旧新往返行为对照，不作为唯一基准

## 约束

- 此基线不改动新系统转换逻辑
- 收尾结束后 MD→DOCX 行为对比以此基线为准，不以收尾后自身行为自证

## VIS-028 反向复用边界

`docx-comprehensive-old-artifact-roundtrip-2026-07-15.md` 与
`old_system_docx_comprehensive_roundtrip_semantics.json` 另将
`sample_golden.docx` 作为旧系统生成的综合 DOCX 输入，执行三项目
DOCX→Markdown 生产入口比较并修复两项 current-only 回归。相邻的
`sample_roundtrip.md` 继续作为旧历史输出来源，但因两个当前 reference 仓实跑均在
列表/引用等处与其漂移，不把它当本轮 exact byte oracle；GOLDEN-002 broader real
document parity 仍 open。

## VIS-099 当前参考边界纠正

2026-07-16 在三个当前 worktree 上重新运行同一 `samples/sample.md` 与
`简体中文通用模板.docx` 后，旧 Tk、旧 PySide6 和当前实现都会把
`aliases[0]` 投影为前导 Title `Test File`。本目录 2026-06-12 固化的
`sample_golden.docx` 没有该 Title，因此它仍是有价值的历史二进制基线，
但不能单独证明两个 reference worktree 的当前标题行为。

`tests/golden/test_md_to_docx_old_baseline.py` 现显式传入真实中文模板，
只允许一个精确的前导 `Test File` Title，并按文档顺序严格比较其余可见
正文、标题与表格；不再通过正文排序掩盖内容移位。三项目综合批次、原始
hash、规范化投影和剩余边界见
`markdown-output-comprehensive-batch-parity-2026-07-16.md` 与
`old_system_markdown_output_batch_semantics.json`。这不是替换历史二进制
fixture，也不是 overall GOLDEN-001 PASS。

后续对 fresh reference 输出的复核还确认：标点结尾标题及其紧邻正文应当
保留为一个带两组 run 格式的 Word 标题段落。历史二进制仍把这两段分开，
因此比较器只对 `samples/sample.md` 中这一组精确标题/正文允许一次合并；
标题文本、正文文本、层级、出现次数或顺序有任何变化仍会失败。

## VIS-101 same-line display math 边界纠正

VIS-101 的历史 OOXML/ordered-comparator 检查确认本 `sample_golden.docx`
在两行块公式前保留一个可见的字面 `$$` 段落；VIS-101 的 Word/WPS 物理
PDF 则来自三个 live worktree 重新执行 Markdown→DOCX 后的 A4 输出，并非
直接把本历史二进制作为分页基准。当前 reference 实跑和修复后的 current
都把该源片段投影为 OMML，不应继续显示分隔符。Comparator 因此只允许从
历史正文序列删除一个精确且唯一的 `$$`，并继续按顺序严格比较全部其他
正文。重复 `$$`、不同文本或正文重排仍失败；这不是替换 fixture 或放宽
一般公式差异。Fresh-output 物理证据见
`markdown-output-physical-word-wps-rendering-2026-07-16.md`。

## VIS-2026-07-17-122 历史二进制物理 oracle 边界

VIS-122 直接检查本文件的 `sectPr`，发现页面为 `1875 × 2652 twips`
（约 `1.302 × 1.842 in`），四边页边距却各为 `1440 twips`，可用宽高为
`-1005/-228 twips`。Word 因此把源投影为 1516 个微型页面；三项目
DOCX→DOC/RTF/ODT 虽 9/9 生产转换成功且同目标 Word 对象计数一致，但 DOC
三份共同无法导出 PDF，RTF/ODT 的 1513/1324 页也没有可接受的版面含义。

本文件继续作为 VIS-028 的结构/正文/标题/表格/notes/OMML/超链接语义
fixture，不能作为物理页面、分页、页边距或 DOC/RTF/ODT broad rendering
oracle。应使用 VIS-101/102 的 live A4 产物或另一个通过页面几何预检的
来源做物理验收。详见
`historical-comprehensive-docx-physical-oracle-boundary-2026-07-17.md`。
