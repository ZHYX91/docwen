# Delimited spreadsheet fidelity

CSV and TSV imports share the same XLSX builder. Every field is literal text, preserving leading zeros, formula-looking text and embedded newlines. Encoding retries rebuild the workbook from scratch.

Before assigning a cell, the builder rejects text longer than 32767 Unicode code points (Python `len`, matching the openpyxl truncation boundary). This is not a byte count or a grapheme count: a supplementary-plane character counts as one and combining characters count separately. The diagnostic contains the source row, column, original length and limit, never the cell contents. CSV, TSV and the XLSX hub route use stable error codes and publish no partial result for this failure. No truncation option or automatic splitting is offered.

Regression tests save and reopen the accepted 32766/32767 boundaries for ASCII, CJK, supplementary characters and multiline fields; 32768/40000 are rejected. This proves serialized XLSX roundtrip with the selected library; real Office interoperability remains a separate candidate acceptance boundary.
