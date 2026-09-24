# Delimited spreadsheet fidelity

CSV and TSV imports share the same XLSX builder. Every field is literal text, preserving leading zeros, formula-looking text and embedded newlines. Encoding retries rebuild the workbook from scratch.

Before assigning a cell, the builder rejects text longer than 32767 UTF-16 code units, matching the native Excel text boundary. This is not a UTF-8 byte count or a grapheme count: a supplementary-plane character counts as two and combining characters count separately. Counting Python code points alone admits files that openpyxl can reopen but Excel refuses to open. The diagnostic contains the source row, column, original UTF-16 length and limit, never the cell contents. CSV, TSV and the XLSX hub route use stable error codes and publish no partial result for this failure. No truncation option or automatic splitting is offered.

Regression tests save and reopen the accepted 32766/32767 UTF-16 boundaries for ASCII, CJK, supplementary characters (including mixed BMP text) and multiline fields; 32768/40000 are rejected. This proves serialized XLSX roundtrip with the selected library; real Office interoperability remains a separate candidate acceptance boundary.

## Formula cache diagnostics

XLSX exports read existing cached values and never recalculate formulas or refresh external links. A formula whose cached scalar is unavailable is exported as an empty field with a warning, aggregated across sheets. The warning retains the total count and at most 20 sheet/cell locations. Existing values, including zero, and ordinary blank cells do not warn. An OOXML formula explicitly carrying `t="str"` and an empty `<v>` is a valid empty string; missing `<v>` or an empty numeric cache remains unavailable. This distinction is checked against package XML because openpyxl maps both cases to `None`.

These warnings propagate through the ordinary conversion result, CLI JSON warnings and text stderr. Cached values are not claimed to be fresh. The source workbook is not modified. Package/GUI display and installed Office checks remain separate acceptance layers.
