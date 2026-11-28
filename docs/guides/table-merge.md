# Merge spreadsheets

[简体中文](table-merge.zh-CN.md)

## Choose the inputs

1. Switch to **Batch** and add the spreadsheets. Use **Settings → Incoming spreadsheets** to choose a default merge mode, or choose a mode beside **Merge Tables** for this task.
2. Select the reference file in the spreadsheet category. Check the **Reference table** label and the collection-file count before starting. The other files in that category are collected into the reference; the inputs are not edited.
3. Only each file's **active worksheet** participates. Activate the intended sheet in your spreadsheet application and save it first. Other sheets in the reference file remain in the result as saved values. Other sheets in collection files are ignored.
4. Check the mode below, start **Merge Tables**, then open the result and its task details. The output is an XLSX workbook. Merged cells in the active sheets are expanded and filled with their top-left value.

## Choose a mode

| Mode | Result |
| --- | --- |
| By row | Compare each incoming row with the reference. Skip an equal row or one whose values are already contained in the reference. Replace only if the incoming row contains all reference values and adds values. Otherwise insert it near the most similar row. |
| By column | Apply the same rule to columns. |
| By cell | After alignment, ignore empty incoming cells, keep a nonempty value when the other side is empty, add numbers (including numeric text), keep equal text once, and join other text with a comma. This mode adds values; it is not a record-deduplication mode. |

Alignment is a heuristic based on matching stored cell values, not a primary-key join. Text stays text: `001`, `1`, `1e0` and numeric `1` are different identities in row/column matching. A real Excel error such as `#N/A` also differs from literal text `#N/A`; coverage, alignment and similarity retain that distinction. Meaningful spaces in text are retained; text containing only whitespace is considered empty. Two complementary records, such as `Alice, 7, blank` and `Alice, blank, 9`, remain separate so neither unique value disappears. Negative alignment offsets make room for leading fields instead of discarding them.

## Formulas and result checks

DocWen reads saved formula results; it does not calculate formulas. A missing result in an active collection sheet or any retained reference sheet stops the merge before an output is produced. The error lists the source file, count and up to 20 cell locations. Recalculate and save the source in Excel, LibreOffice or another suitable spreadsheet application, then retry. A valid cached empty string is treated as empty. A cache being present does **not** prove it is current.

The completion summary reports **added, replaced, skipped and combined** counts for each collection sheet. Counts refer to rows, columns or cells according to the mode; blank items count as skipped. Details contain up to 20 source positions per decision and the applied alignment offsets. These positions refer to the original collection sheet, because later insertions can move output rows and columns. A successful merge has zero unknown formula values. Inspect the result against the sources when alignment or cell addition may be ambiguous.

## Command line

From a source checkout with dependencies synchronized:

```bash
uv run --frozen docwen merge tables base.xlsx collect.xlsx --mode row --output merged.xlsx
```

For a downloaded package, replace `uv run --frozen docwen` with `DocWenCLI.exe` on Windows or `./DocWenCLI` on Linux/macOS. Input order matters: the first file is the reference. Modes are `row`, `col` and `cell`. Use a new output path, or select the overwrite behavior explicitly. See [CLI reference](../cli.md).
