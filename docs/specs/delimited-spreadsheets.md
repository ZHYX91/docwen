# Delimited spreadsheet fidelity

CSV and TSV imports share the same XLSX builder. Every field is literal text, preserving leading zeros, formula-looking text and embedded newlines. Encoding retries rebuild the workbook from scratch.

Admission remains content-first. The generic text sniffer still requires multiple effective delimited records. When a filename explicitly declares `.csv` or `.tsv` and content otherwise remains neutral text, Core may additionally accept exactly one effective multi-column logical record if strict parsing with the declared comma/tab delimiter succeeds. Malformed quoting fails closed, distinctive Markdown or other detected text formats are never overridden, and the existing validated single-column declaration fallback remains separate.

Core admission and spreadsheet conversion share one BOM-first delimited-text decoding contract. UTF-32 LE/BE BOMs are recognized before UTF-16, followed by UTF-8 and UTF-16 BOMs; a BOM-bearing input is decoded only with its declared Unicode family and never retried as another encoding. BOM-less input retries UTF-8 and then GBK. Full-file conversion retries rebuild the workbook from scratch, so a failed candidate cannot leak partial rows into the published XLSX.

Before assigning a cell, the builder rejects text longer than 32767 UTF-16 code units, matching the native Excel text boundary. This is not a UTF-8 byte count or a grapheme count: a supplementary-plane character counts as two and combining characters count separately. Counting Python code points alone admits files that openpyxl can reopen but Excel refuses to open. The diagnostic contains the source row, column, original UTF-16 length and limit, never the cell contents. CSV, TSV and the XLSX hub route use stable error codes and publish no partial result for this failure. No truncation option or automatic splitting is offered.

Regression tests save and reopen the accepted 32766/32767 UTF-16 boundaries for ASCII, CJK, supplementary characters (including mixed BMP text) and multiline fields; 32768/40000 are rejected. This proves serialized XLSX roundtrip with the selected library; real Office interoperability remains a separate candidate acceptance boundary.

Markdown-to-XLSX uses the same 32767 UTF-16 text limit before assigning ordinary table cells or template
YAML, row and column replacements. An oversized final value fails with `MD2XLSX-CELL-TEXT-TOO-LONG` and
a sheet/row/column location; no partial XLSX is published. Plain Markdown-to-CSV does not impose Excel's
cell limit. An XLSX-template-to-CSV path still passes through the bounded XLSX workbook and fails rather
than truncating an oversized replacement.

Table merges validate final string values against the same UTF-16 limit before assignment, including
text accumulated from multiple inputs in cell-merge mode. An oversized value fails with
`MERGE-XLSX-CELL-TEXT-TOO-LONG`, the target sheet/cell and the input names processed so far; the
diagnostic never includes the cell contents. Runtime publishes no XLSX from the failed merge and
preserves existing outputs and source workbooks. Accepted values retain the existing numeric-addition,
text-deduplication and concatenation rules.

## Formula cache diagnostics

XLSX exports read existing cached values and never recalculate formulas or refresh external links. A formula whose cached scalar is unavailable is exported as an empty field with a warning, aggregated across sheets. The warning retains the total count and at most 20 sheet/cell locations. Existing values, including zero, and ordinary blank cells do not warn. An OOXML formula explicitly carrying `t="str"` and an empty `<v>` is a valid empty string; missing `<v>` or an empty numeric cache remains unavailable. This distinction is checked against package XML because openpyxl maps both cases to `None`.

These warnings propagate through the ordinary conversion result, CLI JSON warnings and text stderr. Cached values are not claimed to be fresh. The source workbook is not modified. Package/GUI display and installed Office checks remain separate acceptance layers.

## Cancellation and performance observations

Parsing and export check cancellation in row batches and within wide rows. Formula XML inspection is also cancellable. XLSX serialization is checked immediately before and after save; the library save itself remains synchronous and is not claimed to be interruptible. Cancellation propagates to the runtime as cancellation, closes workbook views, and does not register the current partial artifact. Workspace cleanup owns intermediate files; the existing finalizer cancellation/commit contract decides whether publication has already completed.

The opt-in benchmark covers CSV and TSV roundtrips for 50000×8 cells, 2×16000 cells and 64×4 near-limit text cells. It records parse/save/export elapsed time, peak Python-managed allocation (not total process/native RSS), and observed cancellation response. There are no machine-dependent time thresholds. Run through `tools/qa.py` with `DOCWEN_RUN_SPREADSHEET_BENCHMARK=1`, `DOCWEN_PYTEST_XDIST=0` and `PYTEST_ADDOPTS="packages/plugins/spreadsheet/tests/test_spreadsheet_benchmark.py -s"`; select `--suite full` and a new `--report-output` below workspace acceptance. Preserve the emitted `SPREADSHEET_BASELINE` JSON together with the source identity. Inputs and outputs live only in the QA-owned lease and are removed on success.

Formula-cache validation visits stored worksheet cells only, skips chart sheets, and respects the caller’s worksheet selection. Its memory use does not scale with the rectangular area implied by distant populated cells. The retained base worksheets in table merging use the same check; inactive collection worksheets are outside the merge input scope. Merged values remain saved scalars: literal strings beginning with `=` or matching an Excel error token stay strings, while copied or retained error scalars keep their error type. Row/column coverage, alignment and similarity compare identities that retain the saved error type: a real Excel error and text with the same spelling are distinct. Cell-mode combinations still use the documented addition or text-concatenation policy.


## Delimiter and worksheet identity / 分隔符与工作表身份

Content admission, direct CSV/TSV conversion, the XLSX hub and Markdown export use the same content-proven delimiter. Semicolon-separated multi-record content remains a two-dimensional string matrix. The separately validated comma/tab declaration fallback for neutral single-record input is unchanged. Full-file UTF-8→GBK retries discard the rejected candidate's partial rows.

准入、CSV/TSV 转换、XLSX hub 与 Markdown 导出使用同一内容分隔符事实；多记录分号输入保持二维字符串矩阵。中性单记录继续使用独立验证的 comma/tab 声明规则。完整文件 UTF-8→GBK 重试不会保留失败候选的半成品行。

Markdown/CSV/TSV share the ordered, typed ordinary-worksheet inventory. Chart sheets do not consume worksheet ordinals, count or preferred-output identity. A chart-only workbook fails with `SPREADSHEET-NO-WORKSHEETS` and publishes no empty success. Native cell text crosses one literal Markdown table-cell encoding boundary before generated image/link/OCR fragments are attached. Generated fragments retain their own syntax; source backslash/pipe, punctuation and inline-looking text remain literal values.

Markdown/CSV/TSV 共用有序 typed 普通 worksheet 库存；chartsheet 不占普通工作表的序号、计数或首选产物身份。只有 chartsheet 时以 `SPREADSHEET-NO-WORKSHEETS` 失败，不发布空成功。普通源单元格先经过一次字面 table-cell 编码，再附加生成的图片／链接／OCR 片段；两者不重复转义，源反斜线、pipe 与类似行内语法的文字仍是字面值。

## Bounded native projection / 有界原生投影

Before openpyxl loads a workbook, typed worksheet XML is scanned for stored-cell count and merge coverage. Each request-wide stage has a 1,000,000-cell budget: stored cells, merge coverage, Markdown component grids and copied output blocks. Markdown plans connected components from actual scalar, merge and image coordinates, retaining internal disconnected islands exactly once in row-major order. A distant single value does not imply an A1-origin rectangle. CSV/TSV retain coordinate rectangles through bounded row buffers, reserve cumulative output area before writing, and never create blank Cell objects merely to traverse that rectangle. Overflow fails with `SPREADSHEET-PROJECTION-LIMIT` before the affected expansion.

openpyxl 加载前先按真实 worksheet XML 核存储单元格与 merge 覆盖。每个请求阶段分别有 1,000,000 格上限：存储格、merge 覆盖、Markdown 连通组件网格及输出块复制。Markdown 只由真实 scalar／merge／图片坐标规划，内部断开岛只输出一次并保持 row-major 顺序，远端单值不推导 A1 起始矩形。CSV/TSV 保留坐标矩阵，但在写入前累计核输出面积并以有界行缓冲输出，不为遍历创建空白 Cell。超限在相应展开前以 `SPREADSHEET-PROJECTION-LIMIT` 失败。

Source-extent metadata retains its existing meaning; it is not the sum of compacted blocks. Batch cancellation checks cover scans, wide rows, grid construction and block copies; both workbook views close on failure/cancellation. These cell budgets are not a whole-process RSS or arbitrary decompression/text-byte limit. Native library load/save calls remain synchronous.

源坐标范围 metadata 保持原含义，不改成紧凑输出块之和。扫描、宽行、网格与块复制分批检查取消，失败／取消均关闭两个 workbook views。上述格数预算不是进程 RSS、任意解压或文本总字节上限；原生库 load/save 仍为同步调用。

## Offline XLSX→ODS cached-cell evidence / 离线缓存格交付证明

External formulas are flattened only on the private backend copy. The preparation retains each cached cell's worksheet name, coordinate, OOXML type and value. Missing caches are rejected; `t="str"` with an existing empty `v` is a known empty string. Supported cached constants are numeric, string and boolean; malformed or unsupported cached types fail explicitly. No recalculation or freshness claim is made.

外链公式只在后端私有副本中降为常量，准备阶段保留工作表名称、坐标、OOXML 类型和值。缺失缓存拒绝；`t="str"` 且存在空 `v` 是已知空串。支持 numeric、string、boolean 常量，损坏或不支持的缓存类型明确失败，不声称重算或缓存新鲜。

The final ODS is streamed through its actual document/body/spreadsheet/table/row/cell ownership. Row groups and row/column repeats preserve coordinates without materializing a grid. A matching value elsewhere, on another sheet, or in one cell standing in for two expected cells is insufficient. Numeric equality uses finite Decimal values; booleans remain booleans and strings remain strings. A known empty string may map to a typed empty string or a physically represented untyped blank with no conflicting scalar/formula; an absent cell, covered cell, numeric zero or boolean false is not that mapping. Live external links and the existing protection-loss checks remain enforced.

最终 ODS 按真实 document/body/spreadsheet/table/row/cell 所有权流式核验，行组和行列 repeats 保持坐标且不物化全网格。别处、别表的同值或一个格补两个预期格均不足以证明交付。数值使用有限 Decimal 等价，boolean 与 string 保留类型。已知空串可映射为 typed 空串或没有冲突 scalar/formula 的实际 untyped 空格；缺失、covered cell、数字 0、boolean false 均不是该映射。原 live external link 与保护移除校验继续生效。使用实际包的自动化检查不代替 Office／LibreOffice 的原生转换验收。
