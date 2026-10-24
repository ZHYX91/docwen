# Merge spreadsheets / 表格汇总

## English

### Choose the inputs

1. Switch to **Batch** and add the spreadsheets. Use **Settings → Incoming spreadsheets** to choose a default merge mode, or choose a mode beside **Merge Tables** for this task.
2. Select the reference file in the spreadsheet category. Check the **Reference table** label and the collection-file count before starting. The other files in that category are collected into the reference; the inputs are not edited.
3. Only each file's **active worksheet** participates. Activate the intended sheet in your spreadsheet application and save it first. Other sheets in the reference file remain in the result as saved values. Other sheets in collection files are ignored.
4. Check the mode below, start **Merge Tables**, then open the result and its task details. The output is an XLSX workbook. Merged cells in the active sheets are expanded and filled with their top-left value.

### Choose a mode

| Mode | Result |
| --- | --- |
| By row | Compare each incoming row with the reference. Skip an equal row or one whose values are already contained in the reference. Replace only if the incoming row contains all reference values and adds values. Otherwise insert it near the most similar row. |
| By column | Apply the same rule to columns. |
| By cell | After alignment, ignore empty incoming cells, keep a nonempty value when the other side is empty, add numbers (including numeric text), keep equal text once, and join other text with a comma. This mode adds values; it is not a record-deduplication mode. |

Alignment is a heuristic based on matching stored cell values, not a primary-key join. Text stays text: `001`, `1`, `1e0` and numeric `1` are different identities in row/column matching. A real Excel error such as `#N/A` also differs from literal text `#N/A`; coverage, alignment and similarity retain that distinction. Meaningful spaces in text are retained; text containing only whitespace is considered empty. Two complementary records, such as `Alice, 7, blank` and `Alice, blank, 9`, remain separate so neither unique value disappears. Negative alignment offsets make room for leading fields instead of discarding them.

### Formulas and result checks

DocWen reads saved formula results; it does not calculate formulas. A missing result in an active collection sheet or any retained reference sheet stops the merge before an output is produced. The error lists the source file, count and up to 20 cell locations. Recalculate and save the source in Excel, LibreOffice or another suitable spreadsheet application, then retry. A valid cached empty string is treated as empty. A cache being present does **not** prove it is current.

The completion summary reports **added, replaced, skipped and combined** counts for each collection sheet. Counts refer to rows, columns or cells according to the mode; blank items count as skipped. Details contain up to 20 source positions per decision and the applied alignment offsets. These positions refer to the original collection sheet, because later insertions can move output rows and columns. A successful merge has zero unknown formula values. Inspect the result against the sources when alignment or cell addition may be ambiguous.

### Command line

From a source checkout with dependencies synchronized:

```bash
uv run --frozen docwen merge tables base.xlsx collect.xlsx --mode row --output merged.xlsx
```

For a downloaded package, replace `uv run --frozen docwen` with `DocWenCLI.exe` on Windows or `./DocWenCLI` on Linux/macOS. Input order matters: the first file is the reference. Modes are `row`, `col` and `cell`. Use a new output path, or select the overwrite behavior explicitly. See [CLI reference](../cli.md).

## 中文

### 选择输入

1. 切换到“批量”并加入表格。可在“设置 → 传入表格”选择默认汇总模式，也可在本次“汇总表格”旁选择模式。
2. 在表格类别中选中基表，开始前核对“基表”提示和采集文件数。当前类别的其他文件会汇入该基表，源文件不会被改写。
3. 只处理各文件的**活动工作表**。请先在电子表格软件中激活目标工作表并保存。基表的其他工作表以已保存的值保留在结果中；采集文件的其他工作表不参与汇总。
4. 按下表选择模式，执行“汇总表格”，完成后打开结果和任务详情。输出为 XLSX；活动工作表中的合并单元格会拆开，并填入左上角单元格的值。

### 选择模式

| 模式 | 处理规则 |
| --- | --- |
| 按行 | 将每个采集行与基表比较。完全相同或全部值已被基表包含时跳过；采集行包含基表全部值并有新增值时才替换；其他情况插入到最相似行附近。 |
| 按列 | 对列应用相同规则。 |
| 按单元格 | 对齐后忽略采集侧空值，一侧为空时保留非空值；数值相加（包含数字文本），相同文本保留一份，其他文本用逗号连接。此模式用于累加值，不用于记录去重。 |

自动对齐根据相同的单元格值推测位置，不是按主键连接。按行或列比较时，文本 `001`、`1`、`1e0` 与数字 `1` 是不同身份；真实 Excel 错误值（如 `#N/A`）与同名文本也不同，覆盖比较、对齐和相似度计算均保留这一区别。文本中的有效空格保留，只有全为空白的文本视为空。`Alice, 7, 空` 与 `Alice, 空, 9` 这样的互补记录分别保留，避免丢失任何一侧的独有值。负向对齐会为前方字段留出位置，不将其丢弃。

### 公式与结果核对

DocWen 读取公式已保存的结果，不计算公式。若活动采集工作表或任何保留的基表工作表中存在缺失的公式结果，汇总会在生成产物前停止。错误信息列出源文件、数量及最多 20 个单元格位置。请在 Excel、LibreOffice 等合适的表格软件中重新计算并保存后重试。有效的空字符串缓存按空值处理；**存在缓存不代表缓存已更新**。

完成摘要按每个采集工作表给出**新增、替换、跳过、合并**数量。计数单位由模式决定，分别是行、列或单元格；空白项计入跳过。任务详情记录每类处理最多 20 个来源位置及实际对齐偏移。位置指原采集工作表，因为后续插入可能移动输出中的行列。成功汇总的未知公式值为零。自动对齐或单元格累加可能有歧义时，请对照源文件核查结果。

### 命令行

在已同步依赖的源码目录中运行：

```bash
uv run --frozen docwen merge tables base.xlsx collect.xlsx --mode row --output merged.xlsx
```

使用成品包时，将 `uv run --frozen docwen` 替换为 Windows 的 `DocWenCLI.exe` 或 Linux/macOS 的 `./DocWenCLI`。第一个输入文件是基表，模式可选 `row`、`col`、`cell`。请选择新的输出路径，或明确指定覆盖策略。更多参数见[命令行参考](../cli.md)。
