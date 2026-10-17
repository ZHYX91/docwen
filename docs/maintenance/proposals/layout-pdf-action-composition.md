# OFD/XPS normalization for PDF actions

PDF merging and splitting retain a strict PDF-only Runtime operator. Application route composition now discovers OFD/XPS→PDF normalization before the existing PDF action. A recognized non-PDF layout input gets a protective snapshot and temporary PDF, validated as a readable PDF before forwarding.

Aggregate merge preserves original file order and original output-root semantics; the temporary owner survives through Runtime and is cleaned after completion/failure/cancel. The finalizer publishes only the PDF action output, with optional intermediate publication controlled by existing configuration. The GUI discovers the composed operation for PDF, OFD and XPS without inventing a new plugin route.

The canonical standalone Runtime PDF action still admits PDF only. Machine v2's exact PDF-merge contract is unchanged; no claim is made that arbitrary OFD/XPS content always converts. Tests cover route composition, ordering, failed admission and ownership.
