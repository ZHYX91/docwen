# Office bridge stage integrity

Markdown→DOC/RTF/ODT/WPS/PDF and Markdown→XLS/ODS use an intermediate DOCX or XLSX stage.

* Diagnostics from the intermediate stage remain attached to success **and** Office bridge failures.
* Intermediate stages report 0–70%, bridge work reports 72–85%, and the Runtime finalizer remains responsible for its own 90%/terminal phase. A substage cannot report 100% before finalization.
* When `output.intermediate_files.save_to_output=true`, the explicitly produced intermediate is listed as an auxiliary final artifact. No staging-directory enumeration, filename guessing, or implicit publication is allowed.
* The output manifest/finalizer handle the primary target and declared auxiliary together. On bridge failure nothing is published.
* WPS output retains its existing DOC-backed bridge behavior; file format authenticity must be separately verified with native Office/WPS hosts before changing that contract.

Acceptance includes stage warning propagation, failure diagnostics, progress monotonicity, retained intermediate and unchanged normal artifact count.
