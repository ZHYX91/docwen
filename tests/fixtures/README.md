# 测试资产说明

本目录只存放需要随仓库长期维护的测试资产。目标不是把所有测试输入都外置，而是让需要复用、需要文件语义、或需要整份结果比对的资产有清晰边界。

`old_system_*` 只记录测试证据的历史来源，不构成旧实现或旧缺陷的兼容承诺。当前行为必须单独裁定；旧缺陷不会因为出现在 golden 中而被保留。

- `golden/current_fa01_complete_matrix_semantics.json`
  — VIS-208 is the 91st Golden. It freezes the 18-output
  Markdown→DOCX/XLSX/CSV accounting, corrected three-row template oracle,
  nine fresh project routes, normalized artifacts, Word/WPS/Excel consumers,
  nine PDFs / 105 pages, nine manual contact-sheet checks and zero process
  delta. The exact `PASS_WITH_USER_ACCEPTED_BOUNDARY` is the missing A5
  freeze pane; content/formula/image/validity failures remain errors. Report:
  `fa01-complete-matrix-artifact-oracle-2026-07-24.md`. Overall parity remains
  **NOT PASSED YET**.

- `golden/current_fa06_best_effort_complete_matrix_semantics.json`
  — VIS-206 is the 90th Golden. It freezes the two source identities, 36-slot
  DOCX↔DOC/RTF/ODT accounting, initial converter-created-revision RED and
  repair, warning/source/container/process predicates, Word object projection,
  39 PDFs / 363 pages, nine manual contact-sheet checks and the exact
  `PASS_WITH_USER_ACCEPTED_BOUNDARY` loss statement. Office/PDF/PNG binaries
  remain external. Report:
  `fa06-complete-matrix-artifact-oracle-2026-07-24.md`. Overall parity remains
  **NOT PASSED YET**.

- `golden/current_policy03_preserved_presentation_payloads_semantics.json`
  — VIS-204 is the 89th Golden. It normalizes the reused VIS-129 source
  identities and current 6/6 oracle for ordered chart semantics, exact linked
  embedded XLSX/MP3/MP4/posters, typed snapshot unavailability, same-basename
  final placement, source immutability and zero new Office process. Generated
  PPTX/output binaries remain external. Report:
  `policy03-preserve-presentation-payloads-2026-07-23.md`.

- `golden/old_system_mhtml_to_markdown_semantics.json::fa12_final_artifact_contract_addendum`
  — VIS-176 freezes the external Chromium W3C MHTML hash, 8,579-word/917-link
  oracle, current pre/post-fix artifact projection, strict reference failures
  and hash-reused VIS-127 PPTX N2. VIS-205 additively reconciles those immutable
  facts with VIS-131 and VIS-204. FA-12 status is `FIXED_AND_VERIFIED`: current N1 is
  strict-green, N2 has 34/34 SmartArt plus explicit hidden-slide policy, and
  POLICY-03 passes 6/6. Reference N1 losses remain legacy defects, not accepted
  loss. Reports: `markup-presentation-real-corpus-final-artifact-parity-2026-07-22.md`
  and `fa12-final-artifact-reconciliation-2026-07-23.md`. Overall parity
  remains **NOT PASSED YET**.

- `golden/old_system_docx_to_markdown_rich_semantics.json::fa02_final_artifact_contract_addendum`
  — VIS-175/189 固定外置 EMA Opuviz tracked-changes N1 的来源哈希、Word/OOXML
  形状、三项目完整 source-instance 投影和当前修复后资源可达性；VIS-190 在
  `fa02-legacy-reference-defect-disposition-2026-07-23.md` 将 unchanged Tk/PySide6
  source omissions 明确记为 legacy reference defects，未接受其 loss。FA-02 为
  `FIXED_AND_VERIFIED`；不增加 Golden 数量，overall 仍 **NOT PASSED YET**。

- `golden/old_system_cli_batch_five_route_production_semantics.json` —
  `FA-04-N1` 的 hash-pinned 三项目真实 CLI 五项 mixed-batch 归一化投影；
  固定 DOCX/XLSX/PNG/PDF/bad 的 `5/5/4/1` 顺序、四份可达 Markdown、
  默认 no-OCR、当前同目录 retained-image 复用和 exit-code disposition。
  报告为 `cli-batch-five-route-final-artifact-parity-2026-07-22.md`；只关闭
  FA-04，不代表其它 final-artifact family 或 overall PASS。Golden 当前为 86。

- `golden/old_system_pptx_to_markdown_semantics.json::smartart_hidden_policy_addendum` — VIS-2026-07-17-131 复用固定 Apache POI `60810.pptx`，记录 PowerPoint/package 的 4 diagrams、34 logical nodes、3 hidden slides，以及 current 34/34 ordered bullets、include-all-source-slides policy 和 metadata count。报告为 `apache-poi-smartart-hidden-slide-policy-2026-07-17.md`；原 nonempty lines 与 19 resource name/hash pairs pre/post exact。只关闭 pinned SmartArt/hidden slice，不代表 broad GOLDEN-021/overall PASS；Golden 数量仍为 85。

- `golden/old_system_pptx_to_markdown_semantics.json::legacy_ppt_addendum` — VIS-2026-07-17-130 固定四份 Apache POI legacy PPT、五份 PowerPoint physical screening、current public 4/4、same-WPS-PPTX downstream 12/12 及旧 reference 0/4 Office-16 visibility boundary。current hub 图片链接无资源的真实缺陷由 shared WorkspaceHandle/HubWorkspaceHandle 修复；报告为 `apache-poi-legacy-ppt-bridge-matrix-2026-07-17.md`。不把旧公共路线失败写成成功，不代表 broad GOLDEN-021/overall PASS；Golden 数量仍为 85。

- `golden/old_system_pptx_to_markdown_semantics.json::chart_audio_video_addendum` — VIS-2026-07-17-129 固定三份 Apache POI chart/audio/video PPTX 的 commit/blob/hash、package/PowerPoint source facts、9/9 三项目 PPTX→Markdown 与三页 physical oracle。报告为 `apache-poi-presentation-chart-audio-video-matrix-2026-07-17.md`；normalized semantic/resource projection exact 且无 current-only regression，但 shared chart semantics/snapshot、MP3 与 MP4/video poster loss 未接受，不代表 source-faithful/broad GOLDEN-021 或 overall PASS。Golden 数量仍为 85。

- `golden/old_system_official_office_scripts_workbook_batch_semantics.json::xls_ods_physical_addendum` — VIS-2026-07-17-128 将同一四份 hash-pinned Microsoft Office Scripts XLSX 扩展到 24/24 三项目 XLS/ODS final artifacts、28/28 Excel object/PDF projections、96 output pages 与四张人工检查 contact sheet。8/8 same-target triples 的 workbook semantics/page text/pixels exact，无 current-only regression。共享 ODS form-control、sheet-name、table-style loss 与 formula rewrite 仍未接受；报告为 `office-scripts-workbook-xls-ods-physical-matrix-2026-07-17.md`，Golden 数量仍为 85，不代表 broad GOLDEN-003/overall PASS。

- `golden/old_system_pptx_to_markdown_semantics.json::real_world_rich_presentation_addendum` — VIS-2026-07-17-127 固定 Apache POI `60810.pptx` 的 commit/blob/hash 与 28-slide/section/notes/media/SmartArt package facts，记录三项目 3/3 PPTX→Markdown、294/294 normalized semantic lines、60/60 ordinary texts、6/6 notes、19/19 exact image bytes 和 PowerPoint 25-page source oracle。报告为 `apache-poi-rich-presentation-physical-matrix-2026-07-17.md`；current JPEG MIME 已修复为 `image/jpeg`，shared hidden-slide inclusion 与 11 SmartArt text omission 未接受，不代表 broad GOLDEN-021/overall PASS。Golden 数量仍为 85。

- `golden/old_system_apache_poi_review_field_header_semantics.json::review_header_physical_addendum` — VIS-2026-07-17-121 复用既有三份 Apache POI header/footer/header-picture/comment-media source identity，记录 27/27 三项目 DOC/RTF/ODT 生产转换、40 个 Word PDF/page、12 张人工检查 contact sheet 和 external hashes。报告为 `apache-poi-review-header-physical-matrix-2026-07-17.md`；12/12 same-target content/markup triples exact，DOC/RTF 保留批注内图片，ODT 三项目共同丢失且仍为未接受 fidelity boundary，不代表 broad review UI/GOLDEN-002/overall PASS。Golden 数量仍为 85。

- `golden/old_system_apache_poi_typed_validation_semantics.json::libreoffice_ods_physical_ui_addendum` — VIS-2026-07-17-119 记录 official isolated LibreOffice Calc 26.2.4.2 对 Tk/old/current ODS 的真实 `B2`/`B6` 逐键校验。三项目均保留 prompt/dropdown 并拒绝非法值；`B2` 保留 source custom error，`B6` 共享 LibreOffice localized generic `无效的值。`。报告为 `apache-poi-typed-validation-libreoffice-ui-2026-07-17.md`；whole-string injection 假阴性被 fail-close 排除，shared B6 error-text fidelity/broad corpus/overall PASS 仍 open；Golden 数量仍为 85。

- `golden/old_system_apache_poi_typed_validation_semantics.json::wps_physical_ui_addendum` — VIS-2026-07-17-118 记录 WPS Spreadsheet 12.1.0.26899 对 Tk/old/current XLS 的真实 `B2` 下拉/提示/非法输入一致性，以及 current `B6` 整数拦截；报告为 `apache-poi-typed-validation-wps-ui-2026-07-17.md`。WPS 打开 ODS 要求在线转换，本轮在授权前取消，故 ODS UI 仍为 unavailable boundary，不是 PASS；Golden 数量仍为 85。

- `golden/old_system_apache_poi_typed_validation_semantics.json::final_artifact_ui_addendum` — VIS-2026-07-17-117 复用 VIS-039 的三份 Apache POI/74-rule source identity，记录三项目 XLS/ODS 18/18 生产转换、Excel 16 的 444/444 source-exact rule probes、Tk/old/current XLS 与 current ODS 的四次真实下拉/非法输入交互及 saved-view projection。报告为 `apache-poi-typed-validation-final-artifact-ui-2026-07-17.md`；共享 ODS sheet rename/view reset/hidden-name mojibake 仍未接受，物理 UI 仅抽样一条规则，不代表 broad GOLDEN-003 或 overall PASS。

- `golden/old_system_apache_poi_ooxml_signature_semantics.json` — VIS-2026-07-17-116 固定十份 Apache POI signed/unsigned DOCX/XLSX/PPTX，记录 OOXML signature part、Microsoft Office 16 的 1/2/0 签名识别、三项目 30/30 Markdown normalized body equality 与 21/21 DOCX/XLSX→PDF page/text/geometry/pixel 边界。报告为 `apache-poi-ooxml-signature-boundary-2026-07-17.md`；无 current-only regression，但三项目均无 trust validation、signature preservation 或 derived-artifact warning，仍为未接受 security UX boundary，不代表 GOLDEN-002/003/004 或 overall PASS。

- `golden/old_system_apache_poi_review_field_header_semantics.json::smartdoc_physical_addendum` 与 `golden/old_system_apache_poi_attachment_revision_semantics.json::smartdoc_physical_addendum` — VIS-2026-07-17-113 复用既有 Apache POI 单一 source identity，记录三份复杂 field/revision/five-attachment DOCX 的 27 次 DOC/RTF/ODT 生产转换、30 个 Word PDF、40 页 Poppler pixel equality、Word object projection、九张人工 contact-sheet 检查和 external evidence hashes；报告为 `apache-poi-smartdoc-physical-matrix-2026-07-17.md`。9/9 same-target 三项目物理输出一致，但 shared DOC field error 与 RTF/ODT reflow/text-order 不是 accepted difference，不代表 attachment opening/review UI/broad GOLDEN-002/009 或 overall PASS。

- `golden/old_system_apache_poi_external_link_protection_semantics.json` — VIS-2026-07-17-115 固定 Apache POI commit 的真实 macro/external-link/protection 工作簿，记录三项目 `.xlsm` 非公开 route contract、Tk/current XLSX→ODS external-link 与无密码 protection 语义相等、current bounded-worker cleanup 改进、Excel 密码保护 SaveAs 拒绝、`UpdateLinks=3` 文件/网络访问边界及 15-file external projection。报告为 `apache-poi-external-link-protection-ods-boundary-2026-07-17.md`；`#REF!`/absolute-link、passworded ODS、error taxonomy、broader targets/backends 与 GOLDEN-003/overall PASS 仍 open。

## 当前目录

- `files/`：可复用的输入样例文件。当前用于需要真实路径和文件探测语义的测试，例如 `sample.md`。
- `golden/`：人工可读或可用语义比较工具审阅的期望输出样例。当前锁定 CLI `--json` 输出的稳定结构与代表分支，也包含旧系统 MD→DOCX 二进制/往返基线 `golden/md_to_docx_old/`。

结构化报告资产不放在这里。`skip_report.json`、`not_collected_report.json`、`slow_report.json`、`subprocess_report.json` 与 `missing_marker_report.json` 由 pytest hooks 在运行结束后生成；governed QA 将它们写入仓外 runtime root，属于运行时可见性产物，不属于 fixtures，也不属于 golden。

## 什么时候放进 `files/`

满足以下任一情况时，优先放入 `files/`：

- 测试必须读取真实文件路径、扩展名或文件内容，而不是只消费字符串。
- 样例会被多个测试复用，重复内联会降低可读性。
- 内容较长、包含代码块或接近真实输入，内联后会让测试主体失焦。

以下情况不要强行放进 `files/`：

- 只在单个测试里使用的短字符串。
- 直接内联更容易看懂意图的极小样例。
- 带有无关噪音的大文件或真实用户数据。

原则是“最小、可读、可复用”，而不是“所有输入都文件化”。

## 什么时候使用 `golden/`

满足以下条件时，优先考虑 golden：

- 需要比对整份结构化输出，而不是单个字段。
- 期望结果人工审阅 diff 的成本低，修改后能明确看出行为变化。
- 输出本身属于稳定契约或高价值回归信号，例如 CLI JSON envelope。
- 输出是高价值旧系统二进制 artifact，但已有语义比较工具可把变化解释为段落/标题/表格等可审阅差异，例如 `golden/md_to_docx_old/sample_golden.docx`。
- 输出是高价值旧系统行为或 current enhancement quality 的归一化语义 smoke，且原始 artifact 文件名、路径或 metadata 含时间戳/环境差异，不适合直接提交二进制，例如 `golden/old_system_cli_batch_mixed_invalid_stream_contract.json`、`golden/old_system_cli_batch_mixed_production_smoke_semantics.json`、`golden/old_system_cli_batch_image_production_smoke_semantics.json`、`golden/old_system_cli_batch_layout_image_production_smoke_semantics.json`、`golden/old_system_heic_preprocess_smoke_semantics.json`、`golden/markdown_xlsx_old_system_smoke_semantics.json`（含 MD→XLSX/CSV runtime finalizer 输出门禁）、`golden/old_system_md_xlsx_release_gate_semantics.json`、`golden/old_system_smartdoc_docx_doc_real_smoke_semantics.json`、`golden/old_system_smartdoc_docx_rtf_real_smoke_semantics.json`、`golden/old_system_smartdoc_docx_odt_real_smoke_semantics.json`、`golden/old_system_smartdoc_docx_wps_unsupported_contract.json`、`golden/current_smartdoc_docx_wps_roundtrip_quality_semantics.json`、`golden/current_smartdoc_docx_wps_rich_roundtrip_fidelity_semantics.json`、`golden/old_system_smartdoc_doc_docx_real_smoke_semantics.json`、`golden/old_system_smartdoc_odt_docx_real_smoke_semantics.json`、`golden/old_system_smartdoc_rtf_docx_real_smoke_semantics.json`、`golden/old_system_smartdoc_wps_docx_real_smoke_semantics.json`、`golden/old_system_smartsheet_xlsx_xls_real_smoke_semantics.json`、`golden/old_system_smartsheet_xlsx_xls_style_real_smoke_semantics.json`、`golden/old_system_smartsheet_xlsx_ods_real_smoke_semantics.json`、`golden/old_system_smartsheet_xlsx_ods_style_real_smoke_semantics.json`、`golden/old_system_smartsheet_ods_xlsx_real_smoke_semantics.json`、`golden/old_system_smartsheet_ods_xlsx_style_real_smoke_semantics.json`、`golden/old_system_smartsheet_xls_xlsx_real_smoke_semantics.json`、`golden/old_system_smartsheet_xls_xlsx_style_real_smoke_semantics.json`、`golden/old_system_smartsheet_xlsx_et_real_smoke_semantics.json`、`golden/old_system_smartsheet_xlsx_et_style_real_smoke_semantics.json`、`golden/old_system_smartsheet_et_xlsx_real_smoke_semantics.json`、`golden/old_system_smartsheet_et_xlsx_style_real_smoke_semantics.json`、`golden/old_system_docx_to_markdown_semantics.json`（含 primary Markdown runtime finalizer 输出门禁）、`golden/old_system_docx_to_markdown_rich_semantics.json`（含 tracked insertion/deletion/fldSimple XML boundary）、`golden/old_system_xlsx_to_markdown_semantics.json`（含 primary Markdown runtime finalizer 输出门禁）、`golden/old_system_epub_to_markdown_semantics.json`（含 `image_md` OCR sidecar runtime finalizer 输出门禁）、`golden/old_system_enex_to_markdown_semantics.json`、`golden/old_system_html_to_markdown_semantics.json`、`golden/old_system_mhtml_to_markdown_semantics.json`（含 `image_md` OCR sidecar runtime finalizer 输出门禁）、`golden/old_system_pptx_to_markdown_semantics.json`、`golden/old_system_gongwen_semantics.json`（含主/附件 Markdown runtime finalizer 输出门禁）、`golden/old_system_invoice_cn_semantics.json`（含 PDF/OFD runtime finalizer 输出门禁与 localized title/fixed Chinese invoice schema 边界）、`golden/old_system_i18n_yaml_keys_semantics.json`（含 document/spreadsheet/image/presentation/markup/layout/invoice_cn locale runtime finalizer 输出门禁）、`golden/old_system_image_format_semantics.json`（含 JPG/PNG/WebP runtime finalizer 输出门禁）、`golden/old_system_image_to_markdown_semantics.json`（含 file/base64/image_md runtime finalizer 输出门禁）、`golden/old_system_image_to_pdf_semantics.json`（含 generated EXIF JPEG embedding boundary 与 runtime finalizer 输出门禁）、`golden/old_system_merge_images_to_tiff_semantics.json`（含 RGB/all-RGBA runtime finalizer 输出门禁）、`golden/old_system_layout_pdf_semantics.json`（含 PDF passthrough runtime finalizer 输出门禁、OFD/XPS preprocess parity）、`golden/old_system_pdf_operations_semantics.json`（含 merge/custom/every_page/odd_even split runtime finalizer 输出门禁）、`golden/old_system_merge_tables_semantics.json`（含合并单元格预处理、公式/收集表样式 value-only baseline 与 `_001` finalizer 输出门禁）、`golden/old_system_merge_tables_broader_workbook_semantics.json`（含 mixed sheets、protected sheet、richer merged-cell、base style retention 与 formula-object boundary 投影）、`golden/old_system_proofread_semantics.json`（含三项目 DOCX carrier body paragraphs only 边界与 Markdown sanitizer 旧 `md_spell/text_slicer` 语义）与 `golden/old_system_md_numbering_semantics.json`（含 fenced code block skip、unknown scheme fallback、focused built-in `legal_standard` / `hierarchical_h2_start` scheme matrix、focused custom/user-editable roman-letter scheme projection、focused malformed invalid-placeholder projection 与当前 runtime finalizer 输出门禁）。

- `golden/old_system_smartdoc_rich_outbound_fidelity_semantics.json` 记录同一 rich DOCX 经旧 Tk、旧 PySide6、当前真实 DOCX→DOC/RTF/ODT 后的 normalized roundtrip projection，覆盖样式、comment、field/link、revision policy、图片、页眉页脚与分节，并保留共用 Word ODT numeric `xml:id` strict-parse boundary，不能当作 overall ODT PASS。
- `golden/old_system_smartdoc_odt_rtf_two_hop_semantics.json` 记录同一 rich ODT 经旧 Tk、旧 PySide6、当前真实 ODT→DOCX→RTF 后的 normalized roundtrip projection，并固定当前 private hub、source-owned final name、endpoint metadata、双段 backend 与 final byte metrics；方法、环境与边界见 `smartdoc-odt-rtf-two-hop-artifact-2026-07-14.md`，不代表 all-20-route 或 LibreOffice PASS。
- `golden/old_system_smartdoc_rtf_odt_two_hop_semantics.json` 记录同一 rich RTF 经旧 Tk、旧 PySide6、当前真实 RTF→DOCX→ODT 后的 normalized Word-readback 与 direct ODT package projection，并固定当前 private hub、source-owned final name、endpoint metadata、双段 backend 与 final byte metrics；方法、环境与边界见 `smartdoc-rtf-odt-two-hop-artifact-2026-07-14.md`，不代表 overall ODT、all-20-route 或 LibreOffice PASS。
- `golden/old_system_markdown_output_batch_semantics.json` 记录 VIS-099 同一综合 `samples/sample.md` 经旧 Tk、旧 PySide6、当前真实 MD→DOCX/XLSX/CSV 的紧凑规范化投影、原始 artifact provenance、十一份本地化 DOCX 模板契约和当前 finalizer 落点；方法与边界见 `markdown-output-comprehensive-batch-parity-2026-07-16.md`。它关闭一组代表性 GOLDEN-001 batch 和三项 current-only 缺陷，不是 OOXML byte oracle、物理 Word/WPS 渲染或 broader corpus/overall parity PASS。
- `golden/markdown_xlsx_old_system_smoke_semantics.json` 另含当前 MD→XLSX/CSV runtime finalizer 输出门禁，证明 focused English-template workbook 与 CSV chain 经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径。
- `golden/old_system_docx_to_markdown_semantics.json` 另含当前 primary Markdown runtime finalizer 输出门禁，证明 normalized DOCX Markdown 经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径。
- GOLDEN-002 fixtures cover `conversion.syntax`, unordered-list marker/indent, list-owned code/quote blocks, adjacent-list continuation, `tblInd` table continuation and SDT-internal continuation. The fixture assertions define the maintained regression boundary.
- `golden/old_system_xlsx_to_markdown_semantics.json` 另含当前 primary Markdown runtime finalizer 输出门禁，并记录旧 Tk、旧 PySide6、当前三项目共同的公式 value-only / cached formula value / embedded-image reachability / multi-image order / multi-sheet image order / base64 image-mode data URI / image suppression/no-image output / OCR sidecar no-image output / multi-image OCR sidecar no-image output / multi-image OCR sidecar with-image output / 样式与保护非标准 Markdown 输出边界、CSV direct / TSV old-service-preconvert-current-direct route 边界、TSV blank-row table-block projection 修复、quoted multiline CSV current-enhancement 边界、CSV GBK/UTF-16/semicolon delimiter same-behavior 边界，以及 focused CSV/TSV output-directory batch 边界，证明 normalized XLSX Markdown 经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径，同时证明当前 `.csv` extension dispatch 对 valid quoted multiline CSV 的成功支持不应退回旧项目 content-detection 失败，并证明 focused CSV/TSV batch 经 current runtime 后落入请求输出目录。
- `golden/old_system_image_format_semantics.json` 另含当前 JPG/PNG/WebP runtime finalizer 输出门禁、generated BMP/GIF/TIF matrix projection、generated grayscale/palette `L`/`P` mode matrix、animated GIF / multipage TIFF multiframe boundary、`limit_size` compression success/failure boundary、generated EXIF orientation=6 raw-pixel boundary 和 `image-exif-metadata-boundary-2026-07-05.md` 记录的 generated EXIF metadata clearing boundary，证明 normalized RGBA PNG format outputs 经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径，证明 RGB PNG→BMP/GIF/TIF 的稳定格式/模式/尺寸语义在三项目间一致，证明 grayscale JPG/PNG 恢复旧/当前 `L` representation parity，并记录 transparent palette WebP 的 current alpha-preservation enhancement；同时记录 GIF 动画帧不保留、current GIF→PNG RGBA 透明安全策略、current multipage TIFF→PNG 拆帧增强与 TIFF→TIF 单帧边界，证明 JPEG/WebP 10KB 限制压缩和 PNG 1KB 超限失败语义在三项目间一致，并证明普通 Image→Image 未自动旋转 EXIF orientation、未保留 selected EXIF metadata 是旧 Tk、旧 PySide6、当前共有边界而非当前回归。
- `golden/old_system_image_extended_mode_semantics.json` 记录 generated CMYK / 1-bit / 16-bit grayscale / float 输入到 PNG/WebP/TIFF 的三项目 normalized projection：修复 current CMYK/float→PNG 相对两个旧项目的真实失败，同时保留 current 1-bit PNG/TIFF、16-bit PNG/TIFF、CMYK TIFF 与 float TIFF native-mode fidelity enhancement；observed bytes 与 WebP hash 只作同环境证据，不作为跨 Pillow 版本契约，仍不是 final Image artifact parity PASS。
- `golden/old_system_image_to_markdown_semantics.json` 另含当前 file-mode、base64 no-OCR 与 image_md OCR sidecar runtime finalizer 输出门禁，证明最终输出目录可接收 primary Markdown、retained image、base64 data URI Markdown 和 OCR sidecar，且不泄漏 workspace/staging 路径；该 fixture 还记录 old keep_images=false 输出空图片内容/current omit comment placeholder 增强边界，以及 EXIF frontmatter 非旧/当前标准 Image→Markdown 输出边界。
- `golden/old_system_proofread_semantics.json` 另含当前 Markdown JSON report runtime finalizer 输出门禁，证明 normalized report shape 经 runtime 后落入用户输出目录。
- `golden/old_system_merge_tables_broader_workbook_semantics.json::final_artifact_contract_addendum` 记录 CONTRACT-2026-07-22-001 FA-10 的冻结 B1/N1 输入哈希、六个最终 XLSX provenance、三项目 exact normalized projections、真实 `.xls` hub preparation、bounded process cleanup、current 回归 guard 与 old PySide6 public CLI file-list reference defect；方法和边界见 `merge-tables-final-artifact-parity-2026-07-22.md`。它扩展既有 GOLDEN-013，不增加 Golden fixture 数量，也不关闭 GUI 大表手感或 overall parity。
- `golden/old_system_md_numbering_semantics.json` 另含 focused built-in `legal_standard` / `hierarchical_h2_start` scheme matrix、focused custom/user-editable roman-letter scheme projection、focused malformed invalid-placeholder projection 和当前 numbered Markdown runtime finalizer 输出门禁，证明 focused remove/add/remove+add、内置 scheme、自定义 scheme 与坏 placeholder scheme 输出经 runtime/processor 后仍匹配旧系统 fixture。
- `golden/old_system_epub_to_markdown_semantics.json` 另含当前 primary Markdown runtime finalizer 输出门禁、EPUB multi-resource readable-link cleanup 与 alt-text wiki-link bugfix boundary，以及 `image_md` OCR sidecar runtime finalizer 输出门禁，证明 normalized EPUB Markdown、retained images 与 OCR sidecar 经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径。
- `golden/old_system_enex_to_markdown_semantics.json` 另含当前 primary Markdown runtime finalizer 输出门禁、ENEX embedded image resource runtime finalizer 输出门禁、ENEX multi-resource readable-link cleanup boundary、ENEX main_md OCR inline runtime finalizer 输出门禁与 ENEX image_md OCR sidecar runtime finalizer 输出门禁，证明 normalized ENEX note Markdown、嵌入图片资源、inline OCR 与 OCR sidecar 经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径。
- `golden/old_system_html_to_markdown_semantics.json` 另含当前 primary Markdown、companion image、data URI image、remote image link-only 与 `keep_images=false` OCR main_md runtime finalizer 输出门禁，证明 normalized HTML Markdown、本地图片 artifact、inline PNG artifact、remote link-only 输出和无图片保留 OCR 输出经 runtime 后落入用户输出目录且不泄漏 workspace/staging/data URI 路径或内容。
- `golden/old_system_html_to_markdown_semantics.json` 另含当前 primary Markdown、companion image、HTML multi-resource readable-link current enhancement、data URI image、remote image link-only 与 `keep_images=false` OCR main_md runtime finalizer 输出门禁，证明 normalized HTML Markdown、图片 artifact 和 OCR 文本经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径。
- `golden/old_system_mhtml_to_markdown_semantics.json` 另含当前 primary Markdown、extracted image、MHTML multi-resource readable-link current enhancement 与 `image_md` OCR sidecar runtime finalizer 输出门禁，证明 normalized MHTML Markdown、图片 artifact 和 OCR sidecar 经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径。
- `golden/old_system_pptx_to_markdown_semantics.json` 另含当前 PowerPoint section/notes focused parity、PPTX multi-image order projection、PPTX multi-image OCR sidecar projection、primary Markdown、embedded image artifact 与 `image_md` OCR sidecar runtime finalizer 输出门禁，证明 normalized PPTX Markdown、嵌入图片 artifact 和 OCR sidecar 经 runtime 后落入用户输出目录且不泄漏 workspace/staging 路径。
- `golden/old_system_gongwen_semantics.json` 另含 gongwen 主 18 字段 metadata 与附件 frontmatter 固定中文公文 schema 边界，以及当前主/附件 Markdown runtime finalizer 输出门禁。
- `golden/old_system_invoice_cn_semantics.json` 另含当前 PDF/OFD runtime finalizer 输出门禁、localized generic title key / fixed Chinese invoice business schema 边界，以及 VIS-386 private masked final-artifact addendum。Addendum 只记录 10 PDF / 4 original OFD / 3 PDF-derived image 的 counts、hashes、51-slot result 和用户接受边界；原始私有 bytes/text/HMAC key 不入 fixture。
- `golden/old_system_i18n_yaml_keys_semantics.json` 另含当前 document DOCX、spreadsheet XLSX、image file-mode、presentation PPTX、markup HTML/MHTML/ENEX/EPUB、layout PDF 与 invoice_cn PDF locale runtime finalizer 输出门禁，证明 localized YAML title label 经 `TaskManager -> OutputFinalizer` 后保留在用户输出目录的最终 Markdown 中且不泄漏 workspace/staging 路径；invoice 业务字段与 gongwen metadata 保持旧系统固定中文 schema。
- `golden/old_system_image_to_pdf_semantics.json` 另含 generated EXIF JPEG embedding boundary，以及当前 PNG original/A4、multipage TIFF original 与 generated EXIF JPEG original Image→PDF runtime finalizer 输出门禁，证明 normalized PDF page/embedded-image semantics 经 runtime 后落入用户输出目录。
- `golden/old_system_merge_images_to_tiff_semantics.json` 另含当前 RGB 与 all-RGBA smart Merge Images→TIFF runtime finalizer 输出门禁，证明 normalized TIFF frame semantics 经 runtime 后落入用户输出目录。
- `golden/old_system_layout_pdf_semantics.json` 另含当前 PDF passthrough runtime finalizer 输出门禁、PDF passthrough metadata 投影、OFD/XPS preprocess parity、真实 generated two-page XPS→PDF/PNG/JPG/TIF/Markdown/DOCX projections，证明 normalized Layout PDF passthrough 经 runtime 后落入用户输出目录，且旧 Tk、旧 PySide6 和当前均保留选定 PDF metadata 字段；其中 XPS probe 已用真实 XPS package 替代 PDF substitute；`layout-real-xps-to-pdf-artifact-2026-07-14.md` 记录 direct PDF 一致投影，`layout-real-xps-to-png-artifact-2026-07-14.md` 记录 PNG pixel/finalizer 一致投影与 shared source naming，`layout-real-xps-to-jpg-tif-artifact-2026-07-14.md` 记录 JPG/TIF decoded RGB/finalizer 一致投影，`layout-real-xps-to-markdown-ocr-artifact-2026-07-14.md` 记录 bitmap content/OCR/full-image retention 与 current extracted-image `_preprocess_xps_<uuid>` 命名泄漏修复，`layout-real-xps-to-docx-artifact-2026-07-14.md` 记录三项目 `pdf2docx` fallback 的 DOCX package/section/anchored-media 一致投影和当前环境 hidden Word 超过 120 秒无进展的 open blocker；真实 OFD、external Word/LibreOffice acceptance、broader XPS routes/samples、dependency UX 与 broader PDF/OCR/Document 矩阵仍限定为真实依赖环境和完整 artifact 矩阵。
- `golden/old_system_pdf_operations_semantics.json` 另含当前 merge/custom/every_page/odd_even split runtime finalizer 输出门禁，`pdf-operations-metadata-boundary-2026-07-05.md` 记录的 focused PDF merge/split metadata boundary（selected title/author/subject/keywords/creator/producer fields 在旧 Tk、旧 PySide6 与当前输出中均被清空），`pdf-operations-rotated-page-boundary-2026-07-05.md` 记录的 focused rotated-page geometry boundary（rotation/rect/mediabox/page text 在旧 Tk、旧 PySide6 与当前输出中均保留），`pdf-operations-interactive-geometry-boundary-2026-07-14.md` 记录的真实三项目异构页面 CropBox/rotation/URI link/Text annotation 投影，以及 `pdf-operations-forms-actions-boundary-2026-07-14.md` 记录的 Text widget/FileAttachment preservation 与 pagewise custom split internal-GOTO shared-loss 边界；这些证据证明 normalized PDF page-group semantics 经 runtime 后落入用户输出目录，且上述行为均无当前 parity regression。
- `golden/old_system_merge_tables_semantics.json` 另含合并单元格预处理、公式/收集表样式 value-only baseline 与 `_001` collision finalizer 输出门禁，证明 focused Merge Tables 输出经 runtime 后落入用户输出目录且碰撞命名、merged-cell 展开和 value-only 边界均匹配旧系统。
- `golden/old_system_merge_tables_broader_workbook_semantics.json` 记录更宽 Merge Tables workbook projection：active sheet 处理、base extra sheet 保留、collect extra sheet 忽略、protected sheet 保留、`A1:B2` 展开填值、base `A1` 样式保留、covered `B1` 不复制样式、generated formula 在三项目共同 `data_only=True` / `cell.value` value-only baseline 下不保留，以及 cached formula value focused probe 中 `15 + 5 -> 20` 且不保留公式文本。
- `golden/old_system_smartsheet_rich_two_hop_matrix_semantics.json` — VIS-024 对同一 rich XLS/ODS/ET source bytes 执行旧 Tk、旧 PySide6、current 六条 non-XLSX SmartSheet 双跳的归一化 workbook/package/runtime/process 投影；证据报告为 `smart-sheet-rich-two-hop-matrix-2026-07-14.md`。该 fixture 关闭 focused 六路，不代表 installed LibreOffice、剩余 CSV↔binary、broader workbook 或 GOLDEN-010 final PASS。

以下情况不适合直接做 golden：

- 输出包含路径、耗时、时间戳、随机值等易波动字段，且测试未先归一化。
- 只需断言一两个关键字段，用普通断言更直接。
- 结果格式仍在频繁试验，尚未形成值得长期维护的稳定契约。

## golden 更新规则

刷新 golden 前，必须先回答“这是预期的契约变化，还是实现漂移”。

允许刷新 golden 的场景：

- 需求或契约被明确调整，且对应测试、规范或 schema 已同步更新。
- 输出新增稳定字段，且该字段确实应进入长期回归面。
- 旧样例不再代表当前长期支持的行为，需要用更合适的稳定样例替换。

不应直接刷新 golden 的场景：

- 只是为了让失败测试重新变绿，却没有先解释行为变化。
- 输出里混入了本该在测试里先归一化的波动字段。
- 变更来自偶发环境差异，而不是产品行为的稳定调整。

更新步骤保持最小化：

1. 先在测试中归一化路径、耗时、时间戳等波动值。
2. 人工审 diff，确认变化与预期契约一致。
3. 若顶层结构、必填字段或公开语义发生变化，同步更新 `docs/specs/` 对应规范。
4. 只更新受影响的 golden 文件，不批量重写无关样例。

## 报告资产边界

- `skip_report.json`：记录已被收集、但在运行阶段被跳过的测试项、阶段与原因。
- `not_collected_report.json`：记录因依赖缺失而未进入测试集合的路径、原因与依赖状态。
- `slow_report.json`：记录超过阈值的慢测试项与耗时。
- `subprocess_report.json`：记录统一 subprocess 调用的命令、耗时与返回码。
- `missing_marker_report.json`：记录缺少主分类 marker 的测试项。
- 生成位置：governed QA 默认位于工作区 `.workspace/temp/p<随机名>/reports/`，运行时通过 `DOCWEN_PYTEST_REPORT_DIR` 访问；直接 pytest 也可用该变量重定向。
- 使用方式：本地排查 pytest 收缩或 CI artifact 复盘时读取，不作为断言输入，不提交到仓库。

## 测试体量治理基线

- 该清单是只减不增的迁移棘轮：既有文件不得增长，新文件不得越线；文件拆分到阈值内或删除后必须同步移除失效条目。
- 不得通过提高阈值、增加新基线或移动到未收集目录来让门禁变绿；拆分时保持 node ID 所代表的行为、marker 和 skip 语义。

## 2026-07-22 FA-08 external image corpus

- `golden/old_system_image_format_semantics.json::fa08_final_artifact_contract_addendum`
  records hashes and normalized results for the four frozen B1/B2/N1/N2 images
  and 60 canonical route slots. No new binary image, PDF or Markdown output is
  distributed in this repository.
- N1 is a CC0 iPhone 14 bus photo; N2 is a CC BY 4.0 iPhone 14 Pro road-sign
  photo. Their source pages, attribution and byte hashes are recorded in
  `image-final-artifact-parity-2026-07-22.md`; binaries remain in the external
  evidence root.
- The addendum is deliberately `pass: false`. It locks the repaired empty-OCR
  `image_md` companion projection and the seven shared, unaccepted orientation,
  EXIF/ICC and phone-MPO PDF render failures. Do not regenerate or substitute
  this corpus after observing results.
- `fa08_delivery_first_closure_addendum` preserves that historical VIS-164
  record while documenting VIS-201's selected current closure: `8/8` affected
  current slots pass orientation normalization, supported EXIF, exact ICC and
  warned all-frame MPO delivery. The exact N1 auxiliary-page score remains a
  user-accepted M-B boundary; no external binary is added to this repository.
  Governing evidence is
  `fa08-delivery-first-source-fidelity-implementation-2026-07-23.md`.

## 2026-07-23 POLICY-01 presence-only warning corpus

- `golden/current_policy01_presence_warning_semantics.json` records the
  selected `POLICY-01=B` closure for signed, deterministic invalid and unsigned
  DOCX/XLSX/PPTX.
- The exact Apache POI sources are reused from VIS-116 and are not distributed
  again. The three invalid siblings, 15 delivered Markdown/PDF artifacts and
  CLI/GUI evidence remain under the external VIS-202 evidence root.
- The checked-in fixture normalizes source/invalid hashes, typed warning codes,
  15/15 conversion, 9/9 inspect and 9/9 GUI results. It explicitly records that
  presence-only detection cannot distinguish tampering and does not establish
  signer, integrity, trust, timestamp or revocation truth.
- Governing evidence is
  `policy01-presence-warning-implementation-2026-07-23.md`; overall parity is
  still not passed.

## 2026-07-23 POLICY-02 delivery-first XLSX-to-ODS corpus

- `golden/current_policy02_delivery_first_semantics.json` records the selected
  `POLICY-02 link=B,password=B` closure against the exact VIS-115 linked,
  workbook-password, sheet-password and no-password-control sources.
- The checked-in fixture normalizes source hashes, request/diagnostic
  contracts, Excel 11/11, CLI 3/3, GUI 1/1, cached-value/link/protection
  projections, source immutability and credential redaction. LibreOffice is
  unavailable and explicitly not counted.
- Generated XLSX copies and seven ODS artifacts remain under the external
  VIS-203 evidence root and are not distributed here. The user-accepted
  boundary is static links without future update and intentionally
  unprotected published ODS after correct password plus explicit consent.
- Governing evidence is
  `policy02-flatten-unlock-implementation-2026-07-23.md`; POLICY-03 and overall
  parity remain not passed.

## 2026-07-26 FA-07 accepted SmartSheet boundary

- `golden/current_fa07_complete_matrix_semantics.json` is the normalized
  VIS-213 record for the fixed Apache POI B1 and Ofgem N1 sources across all 24
  XLSX→XLS/ODS→XLSX slots.
- The Ofgem source acquisition, immutable identity and exact 6/6
  qualification are governed by
  `fa07-ofgem-real-financial-model-acquisition-2026-07-24.md` and
  `fa07-ofgem-real-financial-model-acquisition-stage-card-2026-07-24.md`.
- The external workbooks, XLS/ODS/XLSX artifacts, PDFs, PNGs and contact
  sheets remain in the maintainer's private audit archive; they are not
  distributed in this repository.
- The fixture records 20/24 successes, exact current artifact identities,
  source immutability, Excel/WPS/LibreOffice consumer boundaries, the
  delivery-first `ODS_FEATURE_FIDELITY_RISK` diagnostic and the severe N1 ODS
  render/data-validation loss.
- VIS-2026-07-26-385 records the user's exact `那就接受` decision after the
  measured impact was restated. The fixture disposition is now
  `PASS_WITH_USER_ACCEPTED_BOUNDARY`; its measurements and Golden count do not
  change. Governing decision evidence is
  `fa07-ods-fidelity-boundary-acceptance-2026-07-26.md`; the artifact evidence
  remains `fa07-complete-matrix-artifact-oracle-2026-07-24.md`. Objective
  failures and arbitrary ODS loss remain outside the accepted boundary, and
  overall parity remains not passed.
  Stage contract is
  `fa07-ods-fidelity-boundary-acceptance-stage-card-2026-07-26.md`; the
  original artifact stage contract remains
  `fa07-complete-matrix-artifact-oracle-stage-card-2026-07-24.md`.

## 2026-07-22 FA-11 real-corpus fixtures

- `files/proofread_numbering_real/typos-current.toml` and
  `typos-legacy.toml` are syntax-specific file-backed forms of the same real
  `分隔线 <- 分割线` dictionary entry. They are loaded through each project's
  production configuration layer; they are not shipped application defaults.
- `files/proofread_numbering_real/malformed-custom-numbering.md` is the frozen
  N2 file-level corpus. It combines frontmatter, leading H2, skipped levels,
  existing numbering, fenced code, repeated H1, missing templates, invalid
  level references and unknown styles.
- `golden/old_system_proofread_semantics.json::real_dictionary_official_docx_probe`
  records the external official DOCX hash, N1 issue/comment/anchor oracle and
  the old Tk dictionary-consumer defect. The external DOCX and generated
  DOCX/JSON/Markdown artifacts are not distributed in this repository.
- `golden/old_system_md_numbering_semantics.json::broader_malformed_custom_document_probe`
  records the exact N2 output and shared three-project output SHA-256.

## 维护约束

- 新增 `tests/fixtures/` 顶层分类前，先更新本文件，说明用途、边界和更新方式。
- 如果新增 golden 被门禁或规范依赖，同时同步更新 `docs/testing.md` 或对应专题规范。
- 若某类样例只服务单个短测试且不需要真实文件语义，默认继续内联，不为治理而治理。

- `golden/old_system_smartsheet_csv_binary_matrix_semantics.json` — VIS-025 对 CSV→XLS/ODS 与 XLS/ODS/ET→CSV 的同输入三项目 artifact、pre-fix 回归、post-fix sheet/数值投影、current runtime/finalizer 和 process 边界；证据报告为 `smart-sheet-csv-binary-matrix-2026-07-14.md`。该 fixture 使 17 条 SmartSheet route 都有 focused real slice，但不代表 installed LibreOffice、broader/larger inputs、deep features、long-task UX 或 GOLDEN-010 final PASS。
- `golden/old_system_smartdoc_remaining_route_matrix_semantics.json` — VIS-026 对剩余 10 条 SmartDoc route 的 exact source、7 条三项目成功产物、3 条旧系统 unsupported/current WPS enhancement、归一化 Word readback、container/runtime/process 投影；证据报告为 `smartdoc-remaining-route-matrix-2026-07-14.md`。结合既有 fixture 使 20 条 route 都有 focused real evidence，但不代表 installed LibreOffice、broader corpus、deep editable features、long-task UX 或 GOLDEN-009 final PASS。
- `golden/old_system_docx_multilingual_template_batch_semantics.json` — VIS-027 使用三仓共同分发且 byte-identical 的 11 份多语言 DOCX 模板执行 33 次有效 DOCX→Markdown 生产入口转换；锁定 source hash/localized style/Unicode token、三项目 YAML/body 语义、current Unicode 原 stem/finalizer 与排除的未装配 harness 尝试。证据报告为 `docx-multilingual-template-batch-2026-07-15.md`；复杂文档、larger corpus 与 GOLDEN-002 final PASS 仍 open。
- `golden/old_system_docx_comprehensive_roundtrip_semantics.json` — VIS-028 复用 checked-in 旧系统 MD→DOCX 综合制品执行三项目 DOCX→Markdown；锁定 source hash/86 段/34 标题/两表/notes/formulas/rules/link、两项 current-only red→green 修复、old PySide6/current 窄归一化正文、current finalizer 与历史 roundtrip 非 exact oracle 边界。报告为 `docx-comprehensive-old-artifact-roundtrip-2026-07-15.md`；真实业务/公文、resource/review/field、larger corpus 与 GOLDEN-002 final PASS 仍 open。
- `golden/old_system_docx_official_government_list_semantics.json` — VIS-029 使用公开官方政府 DOCX 执行三项目 DOCX→Markdown；锁定 source provenance/hash、81 个 `outlineLvl=9` 正文段、有效 `numId=1`/numbering-disabled `numId=0`、两项 current-only red→green 修复、三项目归一化正文与 current finalizer。报告为 `official-government-docx-list-parity-2026-07-15.md`；外部 binary 不分发，中文列表 format/start 与 broader corpus/resource/review/layout 仍 open。
- `golden/old_system_docx_official_registration_table_semantics.json` — VIS-030 使用公开官方公司登记 DOCX 执行三项目 DOCX→Markdown；锁定 source provenance/hash、20 表/259 行/1,143 positions、两项 current-only red→green 修复、explicit-empty 逐位置 parity、default fill 与 current finalizer。报告为 `official-registration-docx-table-parity-2026-07-15.md`；外部 binary 不分发，broader corpus/field URL/header-footer/resource-review/physical rendering 仍 open。
- `golden/old_system_docx_official_drawingml_textbox_semantics.json` — VIS-031 使用 Cambridge City Council 官方 DrawingML/textbox-dominant DOCX 执行三项目 DOCX→Markdown；锁定 source provenance/hash、15 drawings/32 Choice+Fallback textbox contents/21 unique Choice paragraphs、三项 current-only red→green 修复、current 21/21 source-order projection 与 finalizer。报告为 `official-drawingml-textbox-docx-parity-2026-07-15.md`；外部 binary 不分发，ordinary footer/shape layout/review-field-media/physical rendering 仍 open。
- `golden/old_system_real_world_ocr_quality_semantics.json` — VIS-032 使用 FUNSD 官方 noisy scanned forms 测试集的确定性前五页执行 15 次三项目 English production OCR；锁定 source/page/annotation/model/output hash、预声明 token/character 门槛、三项目 byte-identical 质量投影、current 低置信度过滤修复和代表性 ImagePlugin/runtime/finalizer sidecar。报告为 `real-world-ocr-quality-corpus-2026-07-15.md`；外部 dataset 不分发，handwriting/multilingual/layout-order/photo/scanned-PDF/package/performance breadth 仍 open。
- `golden/old_system_real_world_chinese_photo_ocr_semantics.json` — VIS-033 使用 CCPD2020/CCPD-Green remote ZIP 中预先固定的前五张 road photo 及五个精确 bbox crop 执行 30 次三项目 Chinese production OCR；锁定 source/crop/model/output hash、预声明 similarity 门槛、10/10 byte-identical 输出、失败质量边界和代表性 current ImagePlugin/runtime/finalizer sidecar。报告为 `real-world-chinese-photo-ocr-quality-2026-07-15.md`；外部 archive/photo/crop 不分发，失败门槛不是 accepted difference 或 final PASS。
- `golden/old_system_real_financial_workbook_batch_semantics.json` — VIS-036 使用 Microsoft 官方 Financial Sample 与 ContosoPnL 工作簿执行三项目 XLSX→Markdown 生产入口及 current runtime/finalizer；锁定外部 source hash、700 行 dense table 投影、九 sheet feature-rich 投影、Tk/current 六张 embedded image byte identity与 old PySide6 blank-cell 改进。VIS-107 在同一 fixture 追加 `contoso_pivot_slicer_pdf_addendum`：锁定实际 PivotTable/slicer/data-model 包部件、Excel/WPS/LibreOffice 三项目 90 页渲染、Excel slicer surface、Tk/current WPS equality 与 Print PDF backend priority consumer 修复。报告为 `real-financial-workbook-batch-parity-2026-07-15.md` 和 `contoso-pivot-slicer-pdf-backend-parity-2026-07-17.md`；外部 binary 不分发，editable/interactive PowerPivot/Pivot/slicer、broader print/formula/validation/long-task UX 仍 open，GOLDEN-003/010 未 final PASS。
- `golden/old_system_official_office_scripts_workbook_batch_semantics.json` — VIS-037 使用固定 commit 的四份 Microsoft OfficeDev/Office Scripts 官方 XLSX 执行 12 次三项目 XLSX→Markdown；锁定 source/package/feature inventory、12 个 cached formula、四份 normalized value projection、old PySide6/current raw equality 和代表性 current runtime/finalizer。报告为 `official-office-scripts-workbook-batch-2026-07-15.md`；外部 binary 不分发，静态样本实际为 0 data validation/0 ordinary chart，formula text/missing cache、visual conditional-format/hyperlink/image 与 broader/physical UX 仍 open，GOLDEN-003 未 final PASS。
- `golden/old_system_official_openxml_chart_missing_cache_semantics.json` — VIS-038 筛选固定 `dotnet/Open-XML-SDK` commit 的 109 个 workbook path/96 个唯一 blob，选取真实 `BarChart`/5 组 conditional formatting 与 218 formula/36 missing-cache 工作簿执行 9 次三项目生产转换；锁定两份 normalized value projection、old PySide6/current raw equality、shared chart non-image-extraction 边界与 current runtime/finalizer。报告为 `official-openxml-chart-missing-cache-parity-2026-07-15.md`；32 个 validation 元素均为 prompt-only，typed validation/chart visual/formula evaluation 仍 open，GOLDEN-003 未 final PASS。
- `golden/old_system_apache_poi_typed_validation_semantics.json` — VIS-039 从固定 Apache POI commit 选取三份 typed-validation XLSX，以 9 次三项目 production conversion 锁定 74 条 rule/7 types/7 operators、三份 normalized value projection、old PySide6/current raw equality、source-faithful integer 展示改进与 current runtime/finalizer。报告为 `apache-poi-typed-validation-parity-2026-07-15.md`；Markdown validation preservation/execution 与 Excel dropdown/error UI 仍 open，GOLDEN-003 未 final PASS。
- `golden/old_system_apache_poi_review_field_header_semantics.json` — VIS-041 从同一固定 Apache POI commit 的 128 个 DOCX path/127 unique blob 中 package-screen 四份 distinct complex-field/comment-media/header-footer 文件，以 12 次三项目 production conversion 锁定 equal parsed YAML/body、old Tk/old PySide6 raw equality、saved field/body reachability、shared review/header/media omission 与四份 current runtime/finalizer。报告为 `apache-poi-review-field-header-docx-parity-2026-07-15.md`；shared omission 不是 accepted final fidelity，broader revision/field/attachment/layout/physical UX 与 GOLDEN-002 final PASS 仍 open。
- `golden/old_system_apache_poi_attachment_revision_semantics.json` — VIS-042 从同一固定 Apache POI commit package-screen tracked-revision、single-OLE 与 five-attachment 三份 DOCX，以 9 次 baseline 三项目 conversion 与 3 次 post-fix current finalization 锁定 source/package hashes、revision acceptance、hyperlink/textbox order、shared attachment omission 和 current runtime/finalizer。报告为 `apache-poi-attachment-revision-docx-parity-2026-07-16.md`；current insertion/textbox-anchor gaps 已修，revision metadata、attachment bytes/previews、layout/physical UX 与 GOLDEN-002 final PASS 仍 open。
- `golden/old_system_docx_official_government_list_semantics.json::gongwen_optimizer_probe` — VIS-2026-07-16-093 复用 VIS-029 的单一官方来源事实，实跑三项目 Gongwen 入口并锁定 18 字段、missing-field tolerance、source reachability、current review warning/finalizer metadata 与 set-derived 缺失字段确定性。报告为 `official-government-gongwen-missing-field-parity-2026-07-16.md`；外部 binary 不分发，broader issued-Gongwen/attachment/seal/layout/recipient accuracy 与 GOLDEN-008 final PASS 仍 open。
- `golden/old_system_markdown_multilingual_physical_semantics.json` — VIS-102 固定同一综合 Markdown、三仓 byte-identical 的 11 份本地化 DOCX 模板、33 个真实 DOCX semantic projections、44 个 Word/WPS PDF、246 页 rendered inspection、外部 manifest/projection hash 与 process boundary；报告为 `markdown-output-multilingual-physical-matrix-2026-07-16.md`。Raw DOCX/PDF/PNG 不入库且不是 byte-equality oracle；该 fixture 关闭 all-template Word/WPS physical slice，不代表 broader corpus、installed LibreOffice、其他 final artifacts 或 overall parity PASS。
- `golden/old_system_libreoffice_fallback_matrix_semantics.json` — VIS-103 固定官方隔离 LibreOffice 26.2.4.2、三项目五路/15 次 explicit-empty-COM fallback、五项 normalized artifact projection、45 页 Poppler pixel equality、current filtered-PDF 输出名修复和 live cancellation/process boundary；报告为 `libreoffice-fallback-real-matrix-2026-07-17.md`。该 fixture 关闭 representative production-discoverable LibreOffice slice，不代表 system-wide install、broader corpus、Explorer OLE、其他 final artifacts 或 overall parity PASS。
- `golden/old_system_smartsheet_rich_two_hop_matrix_semantics.json::libreoffice_only_two_hop_addendum` — VIS-2026-07-17-105 在既有 VIS-024 单一事实源中增加官方隔离 LibreOffice 两跳补证据：三项目 XLS→XLSX→ODS、ODS→XLSX→XLS 共 12/12 hops、normalized workbook/package semantics、same-route page pixels、process cleanup 与 external evidence hashes；报告为 `smart-sheet-libreoffice-two-hop-matrix-2026-07-17.md`。共享 conditional-format 丢失和 chart pagination shift 仍是未接受的 fidelity boundary，不代表 broad GOLDEN-010 或 overall PASS。
- `golden/old_system_docx_comprehensive_roundtrip_semantics.json::physical_oracle_boundary` — VIS-2026-07-17-122 直接预检历史 `md_to_docx_old/sample_golden.docx` 的页面几何并执行三项目 DOC/RTF/ODT 共 9 次生产转换。源页约 1.302 × 1.842 英寸而四边边距各 1 英寸，Word 投影为 1516 个微型页面；三项目同目标对象投影一致、无 current-only regression，但该历史二进制只能继续作语义 fixture，不能作分页/物理 oracle。报告为 `historical-comprehensive-docx-physical-oracle-boundary-2026-07-17.md`；VIS-101/102 fresh A4 物理证据不受影响，broad final-artifact parity 仍 open。
- `golden/old_system_docx_official_government_list_semantics.json::real_document_physical_addendum` — VIS-2026-07-17-123 复用同一 hash-pinned 郑州官方规则 DOCX，以有效 A4-like 几何执行三项目 DOC/RTF/ODT 9 次生产转换、10 份 Word/PDF、133 页 120-DPI 渲染和 3 张人工检查 contact sheet。三种 target 的三项目页面/文本/像素均 exact，九份 Word story/paragraph exact to source；DOC near-source，RTF shared line reflow 与 ODT shared 13→14 页扩张不作 accepted source fidelity。报告为 `official-government-docx-physical-matrix-2026-07-17.md`；外部 binary/PDF/PNG 不入库，Golden count 保持 85，broad final-artifact parity 仍 open。
- `golden/old_system_docx_official_registration_table_semantics.json::physical_matrix_addendum` — VIS-2026-07-17-124 将 VIS-030 的同一 73,173-byte 官方无锡登记表扩展为 DOC/RTF/ODT 物理矩阵：13 节/20 表 source 几何有效，9/9 三项目转换、10 份 Word/PDF、150 页与 3 张 contact sheet 证明每个 target 的三项目页数/文本/像素 exact。DOC 对象投影 source-exact；RTF 共享合并说明表并替换四个括号字形，ODT 共享一个空段/空单元格投影及更明显版式变化，后两者均不作 accepted source fidelity。报告为 `official-registration-docx-table-physical-matrix-2026-07-17.md`；外部 binary/PDF/PNG 不入库，Golden count 仍为 85，broad final-artifact parity 仍 open。
- `golden/old_system_docx_to_markdown_rich_semantics.json::physical_matrix_addendum` — VIS-2026-07-17-125 将 checked GOLDEN-002 的同一 38,864-byte 受控 rich probe 扩展为 DOC/RTF/ODT 物理矩阵。9/9 三项目转换、10 份 Word/PDF 与 3 张 contact sheet 证明每个 target 的页数/文本/像素及完整 Word projection exact；image/textbox/nested table/notes/geometry 保留。DOC/RTF 将 OMML 公式变成可见但不可检索图片，ODT 与合成 note body 的标号交互改变首字母，均不作 accepted source fidelity。报告为 `golden002-rich-docx-physical-matrix-2026-07-17.md`；受控 probe 不是官方真实文档，Cambridge exact-source 外部验收仍 open，Golden count 保持 85。
