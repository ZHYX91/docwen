# Delimited spreadsheet fidelity

CSV and TSV imports share the same XLSX builder. Every field is literal text, preserving leading zeros, formula-looking text and embedded newlines. Encoding retries rebuild the workbook from scratch.

Before assigning a cell, the builder rejects text longer than 32767 UTF-16 code units, matching the native Excel text boundary. This is not a UTF-8 byte count or a grapheme count: a supplementary-plane character counts as two and combining characters count separately. Counting Python code points alone admits files that openpyxl can reopen but Excel refuses to open. The diagnostic contains the source row, column, original UTF-16 length and limit, never the cell contents. CSV, TSV and the XLSX hub route use stable error codes and publish no partial result for this failure. No truncation option or automatic splitting is offered.

Regression tests save and reopen the accepted 32766/32767 UTF-16 boundaries for ASCII, CJK, supplementary characters (including mixed BMP text) and multiline fields; 32768/40000 are rejected. This proves serialized XLSX roundtrip with the selected library; real Office interoperability remains a separate candidate acceptance boundary.
