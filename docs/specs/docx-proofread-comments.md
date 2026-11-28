# DOCX proofreading comments

Direct DOCX validation (including Application-preconverted legacy documents) and Markdown-to-DOCX postprocessing use the same validator and comment formatter. Four independent switches control typo rules, sensitive-word rules, symbol correction and symbol pairing. The shipped typo and sensitive-word dictionaries remain empty and user-managed; this does not add grammar correction.

The validator visits main-body paragraphs, table cells and nested tables in document order. Each underlying merged cell is checked once, including horizontal and vertical merges. The existing code/quotation skip options and cancellation checks apply inside tables as well. Paragraph, issue and comment counts include these table paragraphs. The supported scope excludes headers, footers, text boxes, and text inside block-level or inline content controls.

校对按文档顺序遍历主正文段落、表格单元格及嵌套表格。横向或纵向合并后的底层单元格只检查一次。表格内同样执行既有代码／引用跳过选项及取消检查，段落、问题和批注统计包含这些表格段落。页眉、页脚、文本框，以及块级或行内内容控件中的文字不在支持范围内。

Comments use author `DocWen` and initials `DW`. A rule with an explicit `replacement` produces a localized label and `original → replacement`. Sensitive words and unmatched or crossed symbol pairs have no replacement: their comments ask for contextual review and contain no replacement arrow. Stable `source` and `rule_key` identities and Machine fix eligibility are independent of comment language.

The request's `locale` takes precedence, then its frozen `conversion.output_locale`. Display language does not select comment language. DOCX validation capabilities declare `locale` for all twelve shipped languages. Explicit CLI `--lang` is forwarded to locale-aware routes, and Application postprocessing preserves the original request locale. Ordinary profiles default to `zh_CN`; isolated plugin contexts with neither a request locale nor a configured output language retain the English fallback. Markdown report field names and internal rule identifiers are unchanged. See [Task language](task-language.md) for configuration and upgrade behavior.

Pairing uses one nesting stack across symbol types. A closer that matches a non-top opener reports the exact crossing position, then removes that matched opener while leaving inner symbols available for later closers. This finds `([)]` without reporting valid `([])` as an error. Existing contraction, possessive and elision handling and DOCX/Markdown skip policies remain active. The additional reason distinguishes crossed nesting from a genuinely unmatched symbol in DOCX comment wording.

Comment ranges remain half-open ranges in paragraph text. Run splitting and exact range selection are unchanged. Tests write and inspect real DOCX comment parts for every supported locale, four issue kinds, stable author, unchanged input and exact anchors; separate tests cover request/config precedence, CLI and postprocess language transfer, crossed nesting and existing skip/apostrophe behavior. Real Word presentation is an additional host acceptance boundary, not a claim made by those source tests.

Table regressions inspect saved comments and their actual anchors for ordinary, nested and merged cells, preserve the original input bytes, and check Markdown-table conversion through automatic DOCX proofreading before the single final publication.
