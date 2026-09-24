# Delimited spreadsheet fidelity

CSV and TSV imports share the same XLSX builder. Every field is literal text, preserving leading zeros, formula-looking text and embedded newlines. Encoding retries rebuild the workbook from scratch.

Before assigning a cell, the builder rejects text longer than 32767 UTF-16 code units, matching the native Excel text boundary. This is not a UTF-8 byte count or a grapheme count: a supplementary-plane character counts as two and combining characters count separately. Counting Python code points alone admits files that openpyxl can reopen but Excel refuses to open. The diagnostic contains the source row, column, original UTF-16 length and limit, never the cell contents. CSV, TSV and the XLSX hub route use stable error codes and publish no partial result for this failure. No truncation option or automatic splitting is offered.

Regression tests save and reopen the accepted 32766/32767 UTF-16 boundaries for ASCII, CJK, supplementary characters (including mixed BMP text) and multiline fields; 32768/40000 are rejected. This proves serialized XLSX roundtrip with the selected library; real Office interoperability remains a separate candidate acceptance boundary.

## Formula cache diagnostics

XLSX exports read existing cached values and never recalculate formulas or refresh external links. A formula whose cached scalar is unavailable is exported as an empty field with a warning, aggregated across sheets. The warning retains the total count and at most 20 sheet/cell locations. Existing values, including zero, and ordinary blank cells do not warn. An OOXML formula explicitly carrying `t="str"` and an empty `<v>` is a valid empty string; missing `<v>` or an empty numeric cache remains unavailable. This distinction is checked against package XML because openpyxl maps both cases to `None`.

These warnings propagate through the ordinary conversion result, CLI JSON warnings and text stderr. Cached values are not claimed to be fresh. The source workbook is not modified. Package/GUI display and installed Office checks remain separate acceptance layers.

## Cancellation and performance observations

Parsing and export check cancellation in row batches and within wide rows. Formula XML inspection is also cancellable. XLSX serialization is checked immediately before and after save; the library save itself remains synchronous and is not claimed to be interruptible. Cancellation propagates to the runtime as cancellation, closes workbook views, and does not register the current partial artifact. Workspace cleanup owns intermediate files; the existing finalizer cancellation/commit contract decides whether publication has already completed.

The opt-in benchmark covers CSV and TSV roundtrips for 50000×8 cells, 2×16000 cells and 64×4 near-limit text cells. It records parse/save/export elapsed time, peak Python-managed allocation (not total process/native RSS), and observed cancellation response. There are no machine-dependent time thresholds. Run through `tools/qa.py` with `DOCWEN_RUN_SPREADSHEET_BENCHMARK=1`, `DOCWEN_PYTEST_XDIST=0` and `PYTEST_ADDOPTS="packages/plugins/spreadsheet/tests/test_spreadsheet_benchmark.py -s"`; select `--suite full` and a new `--report-output` below workspace acceptance. Preserve the emitted `SPREADSHEET_BASELINE` JSON together with the source identity. Inputs and outputs live only in the QA-owned lease and are removed on success.
